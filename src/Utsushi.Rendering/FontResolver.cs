using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using SkiaSharp;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering;

/// <summary>
/// 帳票が使用するフォント名を <see cref="SKTypeface"/> へ解決する。
/// </summary>
/// <remarks>
/// <para>
/// 実行時に別フォントへ暗黙フォールバックすると帳票の見た目が崩れるため、
/// 既定(<see cref="FontResolverOptions.Strict"/>)では解決できないフォントを例外にする
/// (design.md「Rendering レイヤー」)。
/// </para>
/// <para>
/// 明示的に登録したフォントファイル(<c>.ttf</c>/<c>.otf</c>/<c>.ttc</c>)を最優先で使い、
/// 次に実行環境にインストールされたフォントを探す。
/// </para>
/// </remarks>
public sealed class FontResolver : IDisposable
{
    private readonly FontResolverOptions _options;
    private readonly Dictionary<string, SKTypeface> _registered = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<FontKey, SKTypeface> _cache = new();
    private readonly List<SKTypeface> _owned = new();
    private bool _disposed;

    public FontResolver(FontResolverOptions? options = null)
    {
        _options = options ?? FontResolverOptions.Strict;

        foreach (var (familyName, path) in _options.FontFiles)
        {
            RegisterFontFile(familyName, path);
        }

        foreach (var directory in _options.FontDirectories)
        {
            RegisterFontDirectory(directory);
        }
    }

    /// <summary>フォント名・太字・斜体から書体を解決する。</summary>
    /// <exception cref="FontNotAvailableException">
    /// 厳格モードでフォントが見つからない場合。
    /// </exception>
    public SKTypeface Resolve(FontStyle font)
    {
        ThrowIfDisposed();

        var key = new FontKey(font.Name, font.Bold, font.Italic);
        if (_cache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        var resolved = ResolveCore(font);
        return _cache.GetOrAdd(key, resolved);
    }

    private SKTypeface ResolveCore(FontStyle font)
    {
        var weight = font.Bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
        var slant = font.Italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
        var style = new SKFontStyle(weight, SKFontStyleWidth.Normal, slant);

        if (_registered.TryGetValue(font.Name, out var registered))
        {
            return registered;
        }

        var typeface = SKTypeface.FromFamilyName(font.Name, style);

        // SkiaSharp は未知のファミリ名に対して既定フォントを返すため、
        // 実際に要求した書体が得られたかを名前で確認する。
        if (typeface is not null
            && string.Equals(typeface.FamilyName, font.Name, StringComparison.OrdinalIgnoreCase))
        {
            return typeface;
        }

        typeface?.Dispose();

        if (_options.Mode == FontResolutionMode.Strict)
        {
            throw new FontNotAvailableException(
                font.Name,
                $"フォント '{font.Name}' が実行環境に見つかりません。"
                + "帳票の見た目が崩れるため代替フォントへのフォールバックは行いません。"
                + "フォントをインストールするか、FontResolverOptions でフォントファイルを登録してください。");
        }

        if (_options.FallbackFamilyName is { } fallbackName)
        {
            var fallback = SKTypeface.FromFamilyName(fallbackName, style);
            if (fallback is not null)
            {
                return fallback;
            }
        }

        return SKTypeface.Default;
    }

    /// <summary>フォントファイルをファミリ名に紐づけて登録する。</summary>
    public void RegisterFontFile(string familyName, string path)
    {
        ThrowIfDisposed();

        if (!File.Exists(path))
        {
            throw new FontNotAvailableException(
                familyName, $"登録しようとしたフォントファイルが存在しません: {path}");
        }

        var typeface = SKTypeface.FromFile(path)
            ?? throw new FontNotAvailableException(
                familyName, $"フォントファイルを読み込めません(対応していない形式の可能性があります): {path}");

        _owned.Add(typeface);
        _registered[familyName] = typeface;
    }

    /// <summary>ディレクトリ内のフォントファイルを、その内部のファミリ名で一括登録する。</summary>
    public void RegisterFontDirectory(string directory)
    {
        ThrowIfDisposed();

        if (!Directory.Exists(directory))
        {
            return;
        }

        foreach (var path in Directory.EnumerateFiles(directory))
        {
            var extension = Path.GetExtension(path);
            if (!extension.Equals(".ttf", StringComparison.OrdinalIgnoreCase)
                && !extension.Equals(".otf", StringComparison.OrdinalIgnoreCase)
                && !extension.Equals(".ttc", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var typeface = SKTypeface.FromFile(path);
            if (typeface is null)
            {
                continue;
            }

            _owned.Add(typeface);
            _registered[typeface.FamilyName] = typeface;
        }
    }

    /// <summary>
    /// 指定されたフォントがすべて解決できることを事前に検証する。
    /// </summary>
    /// <remarks>
    /// デプロイ時・起動時にフォント不足を検出するための入口(タスク6.2)。
    /// 実際の変換処理の前に呼び出すことで、描画途中での失敗を避けられる。
    /// </remarks>
    /// <exception cref="FontNotAvailableException">解決できないフォントがある場合。</exception>
    public void EnsureAvailable(IEnumerable<FontStyle> fonts)
    {
        foreach (var font in fonts)
        {
            Resolve(font);
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        foreach (var typeface in _owned)
        {
            typeface.Dispose();
        }

        _owned.Clear();
        _registered.Clear();
        _cache.Clear();
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(FontResolver));
        }
    }

    private readonly record struct FontKey(string Name, bool Bold, bool Italic);
}

/// <summary>フォント解決の方針。</summary>
public enum FontResolutionMode
{
    /// <summary>解決できないフォントを例外にする(既定)。</summary>
    Strict = 0,

    /// <summary>解決できない場合に代替フォントを使う。見た目が崩れる可能性がある。</summary>
    AllowFallback,
}

/// <summary>
/// <see cref="FontResolver"/> の設定。
/// </summary>
/// <param name="Mode">フォント解決の方針。</param>
/// <param name="FontFiles">ファミリ名 → フォントファイルパスの明示的な対応。</param>
/// <param name="FontDirectories">フォントファイルを一括登録するディレクトリ。</param>
/// <param name="FallbackFamilyName">
/// <see cref="FontResolutionMode.AllowFallback"/> のときに使う代替フォント名。
/// </param>
public sealed record FontResolverOptions(
    FontResolutionMode Mode,
    IReadOnlyDictionary<string, string> FontFiles,
    IReadOnlyList<string> FontDirectories,
    string? FallbackFamilyName)
{
    /// <summary>厳格モード(既定)。</summary>
    public static FontResolverOptions Strict { get; } = new(
        FontResolutionMode.Strict,
        new Dictionary<string, string>(),
        Array.Empty<string>(),
        null);

    /// <summary>フォールバックを許容する設定。見た目の再現性より処理の継続を優先する場合に使う。</summary>
    public static FontResolverOptions AllowFallback(string? fallbackFamilyName = null) => new(
        FontResolutionMode.AllowFallback,
        new Dictionary<string, string>(),
        Array.Empty<string>(),
        fallbackFamilyName);
}
