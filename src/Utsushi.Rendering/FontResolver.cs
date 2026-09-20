using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using SkiaSharp;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering
{
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
    /// <para>
    /// 太字・斜体は、実フォントがその字形を持つ場合のみ実字形を使う。持たない場合は
    /// 通常字形の書体を返し、<see cref="ResolvedTypeface.SynthesizeBold"/> /
    /// <see cref="ResolvedTypeface.SynthesizeItalic"/> を立てて描画側に装飾を委ねる。
    /// SkiaSharp に合成させると PDF が Type 3 フォントになり、文字列検索ができなくなるため。
    /// </para>
    /// </remarks>
    public sealed class FontResolver : IDisposable
    {
        /// <summary>フォントの同一性判定に使う先頭バイト数。</summary>
        private const int FontIdentityProbeBytes = 4096;

        private readonly FontResolverOptions _options;
        private readonly Dictionary<FontKey, SKTypeface> _registered = new();
        private readonly ConcurrentDictionary<FontKey, ResolvedTypeface> _cache = new();
        private readonly List<SKTypeface> _owned = new();
        private readonly object _registrationLock = new();
        private bool _disposed;

        public FontResolver(FontResolverOptions? options = null)
        {
            _options = options ?? FontResolverOptions.Strict;

            foreach (var (key, path) in _options.FontFiles)
            {
                RegisterFontFile(key, path);
            }

            foreach (var directory in _options.FontDirectories)
            {
                RegisterFontDirectory(directory);
            }
        }

        /// <summary>
        /// フォント名・太字・斜体から書体を解決する。
        /// </summary>
        /// <exception cref="FontNotAvailableException">
        /// 厳格モードでフォントが見つからない場合。
        /// </exception>
        public ResolvedTypeface Resolve(FontStyle font)
        {
            ThrowIfDisposed();

            var key = new FontKey(font.Name, font.Bold, font.Italic);
            if (_cache.TryGetValue(key, out var cached))
            {
                return cached;
            }

            return _cache.GetOrAdd(key, ResolveCore(font));
        }

        private ResolvedTypeface ResolveCore(FontStyle font)
        {
            // 1. 明示登録されたフォントファイル(スタイル指定つき)を最優先で使う。
            if (TryGetRegistered(font.Name, font.Bold, font.Italic, out var registeredExact))
            {
                return new ResolvedTypeface(registeredExact, SynthesizeBold: false, SynthesizeItalic: false);
            }

            // 2. 明示登録の通常字形があれば、それを使って装飾を合成する。
            if (TryGetRegistered(font.Name, bold: false, italic: false, out var registeredRegular))
            {
                return new ResolvedTypeface(registeredRegular, font.Bold, font.Italic);
            }

            // 3. 実行環境にインストールされたフォントを探す。
            var regular = FindInstalled(font.Name, bold: false, italic: false);
            if (regular is null)
            {
                return HandleMissingFont(font);
            }

            if (!font.Bold && !font.Italic)
            {
                return new ResolvedTypeface(regular, false, false);
            }

            // 太字/斜体が要求された場合、実字形を持つ別ファイルがあるかを確認する。
            // SKTypeface.FromFamilyName は字形が無くても合成した書体を返すため、
            // 通常字形とフォントデータが同一かどうかで「実字形の有無」を判定する。
            var styled = FindInstalled(font.Name, font.Bold, font.Italic);
            if (styled is not null && !IsSameFontData(regular, styled))
            {
                regular.Dispose();
                return new ResolvedTypeface(styled, false, false);
            }

            styled?.Dispose();
            return new ResolvedTypeface(regular, font.Bold, font.Italic);
        }

        private ResolvedTypeface HandleMissingFont(FontStyle font)
        {
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
                var fallback = FindInstalled(fallbackName, bold: false, italic: false);
                if (fallback is not null)
                {
                    return new ResolvedTypeface(fallback, font.Bold, font.Italic);
                }
            }

            return new ResolvedTypeface(SKTypeface.Default, font.Bold, font.Italic);
        }

        /// <summary>
        /// インストール済みフォントから、要求したファミリ名と厳密に一致する書体を探す。
        /// </summary>
        /// <remarks>
        /// SkiaSharp は未知のファミリ名に対しても既定フォントを返すため、
        /// 得られた書体のファミリ名を照合して「本当に要求したフォントか」を確認する。
        /// </remarks>
        private static SKTypeface? FindInstalled(string familyName, bool bold, bool italic)
        {
            var style = new SKFontStyle(
                bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal,
                italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);

            var typeface = SKTypeface.FromFamilyName(familyName, style);
            if (typeface is null)
            {
                return null;
            }

            if (string.Equals(typeface.FamilyName, familyName, StringComparison.OrdinalIgnoreCase))
            {
                return typeface;
            }

            typeface.Dispose();
            return null;
        }

        /// <summary>
        /// 2つの書体が同一のフォントデータかどうかを判定する。
        /// </summary>
        /// <remarks>
        /// 同一なら、要求した太字/斜体の字形をフォントが持たず SkiaSharp が合成したことを意味する。
        /// 全バイトの比較はフォントサイズ(日本語フォントで数MB)を考えると重いため、
        /// バイト長と先頭 <see cref="FontIdentityProbeBytes"/> バイトで判定する。
        /// 先頭にはテーブルディレクトリが含まれ、別ファイルであればまず一致しない。
        /// </remarks>
        private static bool IsSameFontData(SKTypeface left, SKTypeface right)
        {
            using var leftStream = left.OpenStream();
            using var rightStream = right.OpenStream();

            if (leftStream is null || rightStream is null)
            {
                // データを読めない場合は「別物と判断できない」ため、合成扱いにして安全側へ倒す。
                return true;
            }

            if (leftStream.Length != rightStream.Length)
            {
                return false;
            }

            var leftHead = ReadHead(leftStream);
            var rightHead = ReadHead(rightStream);

            if (leftHead.Length != rightHead.Length)
            {
                return false;
            }

            for (var i = 0; i < leftHead.Length; i++)
            {
                if (leftHead[i] != rightHead[i])
                {
                    return false;
                }
            }

            return true;
        }

        private static byte[] ReadHead(SKStreamAsset stream)
        {
            var length = Math.Min(FontIdentityProbeBytes, stream.Length);
            var buffer = new byte[length];
            stream.Read(buffer, buffer.Length);
            return buffer;
        }

        /// <summary>
        /// フォントファイルを、フォントキーに紐づけて登録する。
        /// </summary>
        /// <param name="fontKey">
        /// <c>ファミリ名</c>、または <c>ファミリ名:bold</c> / <c>:italic</c> / <c>:bolditalic</c>。
        /// スタイル指定なしで登録したフォントは、太字/斜体が要求された際に描画側で合成される。
        /// </param>
        /// <param name="path">フォントファイルのパス。</param>
        public void RegisterFontFile(string fontKey, string path)
        {
            ThrowIfDisposed();

            var key = FontKey.Parse(fontKey);

            if (!File.Exists(path))
            {
                throw new FontNotAvailableException(
                    key.Name, $"登録しようとしたフォントファイルが存在しません: {path}");
            }

            var typeface = SKTypeface.FromFile(path)
                ?? throw new FontNotAvailableException(
                    key.Name, $"フォントファイルを読み込めません(対応していない形式の可能性があります): {path}");

            lock (_registrationLock)
            {
                _owned.Add(typeface);
                _registered[key] = typeface;
            }
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

                // ファイル自身が申告するスタイルで登録する。太字ファイルは太字キーに入る。
                var key = new FontKey(typeface.FamilyName, typeface.IsBold, typeface.IsItalic);

                lock (_registrationLock)
                {
                    _owned.Add(typeface);
                    _registered[key] = typeface;
                }
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

        private bool TryGetRegistered(string familyName, bool bold, bool italic, out SKTypeface typeface)
        {
            lock (_registrationLock)
            {
                return _registered.TryGetValue(new FontKey(familyName, bold, italic), out typeface!);
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            lock (_registrationLock)
            {
                foreach (var typeface in _owned)
                {
                    typeface.Dispose();
                }

                _owned.Clear();
                _registered.Clear();
            }

            _cache.Clear();
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(FontResolver));
            }
        }

        /// <summary>ファミリ名とスタイルの組。ファミリ名の大小は区別しない。</summary>
        private readonly struct FontKey : IEquatable<FontKey>
        {
            public FontKey(string name, bool bold, bool italic)
            {
                Name = name;
                Bold = bold;
                Italic = italic;
            }

            public string Name { get; }

            public bool Bold { get; }

            public bool Italic { get; }

            /// <summary>"MS PGothic" / "MS PGothic:bold" 形式のキー文字列を解釈する。</summary>
            public static FontKey Parse(string fontKey)
            {
                if (string.IsNullOrWhiteSpace(fontKey))
                {
                    throw new ArgumentException("フォントキーが空です。", nameof(fontKey));
                }

                var separator = fontKey.LastIndexOf(':');
                if (separator <= 0)
                {
                    return new FontKey(fontKey.Trim(), false, false);
                }

                var name = fontKey.Substring(0, separator).Trim();
                var style = fontKey.Substring(separator + 1).Trim().ToLowerInvariant();

                return style switch
                {
                    "bold" => new FontKey(name, true, false),
                    "italic" => new FontKey(name, false, true),
                    "bolditalic" => new FontKey(name, true, true),
                    "regular" or "" => new FontKey(name, false, false),
                    _ => throw new ArgumentException(
                        $"フォントキーのスタイル '{style}' は不正です(bold / italic / bolditalic / regular)。",
                        nameof(fontKey)),
                };
            }

            public bool Equals(FontKey other) =>
                Bold == other.Bold
                && Italic == other.Italic
                && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase);

            public override bool Equals(object? obj) => obj is FontKey other && Equals(other);

            public override int GetHashCode() =>
                ((StringComparer.OrdinalIgnoreCase.GetHashCode(Name) * 397) ^ Bold.GetHashCode()) * 397
                ^ Italic.GetHashCode();
        }
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
    /// <param name="FontFiles">
    /// フォントキー → フォントファイルパスの明示的な対応。
    /// キーは <c>ファミリ名</c> または <c>ファミリ名:bold</c> / <c>:italic</c> / <c>:bolditalic</c>。
    /// </param>
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
}
