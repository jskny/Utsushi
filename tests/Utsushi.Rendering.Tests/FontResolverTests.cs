using System;
using System.Collections.Generic;
using System.IO;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests;

/// <summary>
/// フォント解決の検証(タスク6.2)。
/// </summary>
public sealed class FontResolverTests
{
    private static FontStyle Font(string name) => FontStyle.Default with { Name = name };

    [Fact]
    public void 厳格モードでは未検出フォントを例外にする()
    {
        using var resolver = new FontResolver(FontResolverOptions.Strict);

        var ex = Assert.Throws<FontNotAvailableException>(
            () => resolver.Resolve(Font("存在しないフォント-ZZZ")));

        Assert.Equal("存在しないフォント-ZZZ", ex.FontName);
        Assert.Equal(ProcessingStage.Rendering, ex.Stage);
    }

    [Fact]
    public void フォールバック許容なら代替フォントを返す()
    {
        using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

        var typeface = resolver.Resolve(Font("存在しないフォント-ZZZ"));

        Assert.NotNull(typeface);
    }

    [Fact]
    public void 事前検証で未検出フォントをまとめて検出できる()
    {
        using var resolver = new FontResolver(FontResolverOptions.Strict);

        var ex = Assert.Throws<FontNotAvailableException>(
            () => resolver.EnsureAvailable(new[] { Font("存在しないフォント-ZZZ") }));

        Assert.Equal("存在しないフォント-ZZZ", ex.FontName);
    }

    [Fact]
    public void 存在しないフォントファイルの登録は例外になる()
    {
        using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

        Assert.Throws<FontNotAvailableException>(
            () => resolver.RegisterFontFile("Any", Path.Combine(Path.GetTempPath(), "no-such-font.ttf")));
    }

    [Fact]
    public void 明示登録したフォントは名前で解決できる()
    {
        var fontPath = FindAnyInstalledFont();
        if (fontPath is null)
        {
            // 実行環境にフォントが1つも無い場合は検証できない。
            return;
        }

        using var resolver = new FontResolver(
            new FontResolverOptions(
                FontResolutionMode.Strict,
                new Dictionary<string, string> { ["社内帳票フォント"] = fontPath },
                Array.Empty<string>(),
                null));

        var typeface = resolver.Resolve(Font("社内帳票フォント"));

        Assert.NotNull(typeface);
    }

    private static string? FindAnyInstalledFont()
    {
        foreach (var directory in new[] { "/usr/share/fonts", "/usr/local/share/fonts" })
        {
            if (!Directory.Exists(directory))
            {
                continue;
            }

            foreach (var path in Directory.EnumerateFiles(directory, "*.ttf", SearchOption.AllDirectories))
            {
                return path;
            }
        }

        return null;
    }
}
