using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests;

/// <summary>
/// PDF描画の検証(要件5.1-5.4、タスク6.1-6.3)。
/// </summary>
public sealed class SkiaPdfRendererTests : IDisposable
{
    private readonly FontResolver _fontResolver;
    private readonly SkiaFontMetricsProvider _metrics;

    public SkiaPdfRendererTests()
    {
        // CI/開発環境に特定フォントがあるとは限らないため、描画テストではフォールバックを許容する。
        _fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
        _metrics = new SkiaFontMetricsProvider(_fontResolver);
    }

    public void Dispose() => _fontResolver.Dispose();

    private static PagedLayout Layout(int pageCount = 1, IReadOnlyList<DrawCommand>? commands = null)
    {
        var pages = new List<PageLayout>();
        for (var i = 1; i <= pageCount; i++)
        {
            pages.Add(new PageLayout(
                PaperSize.A4,
                PageOrientation.Portrait,
                PaperSize.A4.WidthPt,
                PaperSize.A4.HeightPt,
                commands ?? DefaultCommands(),
                i,
                (1, 10),
                (1, 5),
                1.0));
        }

        return new PagedLayout(pages, "test-report", "テストシート");
    }

    private static IReadOnlyList<DrawCommand> DefaultCommands() => new DrawCommand[]
    {
        new FillRectCommand(new RectPt(50, 50, 200, 30), new ArgbColor(0xFF, 0xDC, 0xE6, 0xF1)),
        new LineCommand(new PointPt(50, 50), new PointPt(250, 50), ArgbColor.Black, 0.75, LineDashStyle.Solid),
        new TextCommand(new PointPt(55, 70), "テスト", FontStyle.Default, TextAnchor.Left, null),
    };

    [Fact]
    public void 単一のPDFファイルとして出力される()
    {
        var renderer = new SkiaPdfRenderer(_metrics);
        using var output = new MemoryStream();

        renderer.Render(Layout(pageCount: 3), output);

        var bytes = output.ToArray();
        Assert.True(bytes.Length > 0);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));

        // 3ページぶんの Page オブジェクトが含まれる
        var content = Encoding.Latin1.GetString(bytes);
        Assert.Contains("/Type /Page", content);
    }

    [Fact]
    public void ページサイズは用紙と向きに対応する()
    {
        var renderer = new SkiaPdfRenderer(_metrics);
        using var output = new MemoryStream();

        var landscape = new PagedLayout(
            new[]
            {
                new PageLayout(
                    PaperSize.A4, PageOrientation.Landscape,
                    PaperSize.A4.HeightPt, PaperSize.A4.WidthPt,
                    DefaultCommands(), 1, (1, 1), (1, 1), 1.0),
            },
            "test-report", "テストシート");

        renderer.Render(landscape, output);

        var content = Encoding.Latin1.GetString(output.ToArray());
        var mediaBox = System.Text.RegularExpressions.Regex.Match(
            content, @"/MediaBox\s*\[\s*([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s*\]");

        Assert.True(mediaBox.Success, "PDF に /MediaBox が含まれるはず");

        var width = double.Parse(mediaBox.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture);
        var height = double.Parse(mediaBox.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);

        // A4横 = 841.89 x 595.28 pt。PDF側の数値は有効桁が丸められるため許容誤差を持たせる。
        Assert.Equal(PaperSize.A4.HeightPt, width, precision: 0);
        Assert.Equal(PaperSize.A4.WidthPt, height, precision: 0);
        Assert.True(width > height, "横向きなので幅のほうが大きいはず");
    }

    [Fact]
    public void ページが無ければエラーになる()
    {
        var renderer = new SkiaPdfRenderer(_metrics);
        using var output = new MemoryStream();

        var ex = Assert.Throws<PdfRenderingException>(
            () => renderer.Render(new PagedLayout(Array.Empty<PageLayout>(), "r", "s"), output));

        Assert.Equal(ProcessingStage.Rendering, ex.Stage);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public void アウトライン出力はフォントを埋め込まない()
    {
        using var embedded = new MemoryStream();
        using var outlined = new MemoryStream();

        new SkiaPdfRenderer(_metrics, PdfRenderOptions.Default).Render(Layout(), embedded);
        new SkiaPdfRenderer(_metrics, PdfRenderOptions.OutlineText).Render(Layout(), outlined);

        // フォントを埋め込まないぶん、アウトライン出力のほうが小さくなる。
        Assert.True(
            outlined.Length < embedded.Length,
            $"アウトライン出力のほうが小さいはず (outline={outlined.Length}, embed={embedded.Length})");

        var content = Encoding.Latin1.GetString(outlined.ToArray());
        Assert.DoesNotContain("/FontFile", content);
    }

    [Fact]
    public void ファイル出力は一時ファイル経由で確定する()
    {
        var renderer = new SkiaPdfRenderer(_metrics);
        var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
        var path = Path.Combine(directory, "out.pdf");

        try
        {
            renderer.RenderToFile(Layout(), path);

            Assert.True(File.Exists(path));
            Assert.False(File.Exists(path + ".utsushi-tmp"), "一時ファイルが残ってはいけない");
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (Directory.Exists(directory))
            {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void 描画に失敗しても出力先には何も書き込まない()
    {
        var renderer = new SkiaPdfRenderer(_metrics);
        using var output = new MemoryStream();

        // 未知の描画命令を1つ混ぜて途中で失敗させる。
        var layout = Layout(commands: new DrawCommand[]
        {
            new FillRectCommand(new RectPt(0, 0, 10, 10), ArgbColor.Black),
            new UnknownCommand(),
        });

        Assert.Throws<PdfRenderingException>(() => renderer.Render(layout, output));

        // 要件5.4: 不完全なPDFを出力先に残さない
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public void 既存ファイルがあっても失敗時は上書きされない()
    {
        var renderer = new SkiaPdfRenderer(_metrics);
        var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "existing.pdf");

        try
        {
            File.WriteAllText(path, "既存の内容");

            var layout = Layout(commands: new DrawCommand[] { new UnknownCommand() });
            Assert.Throws<PdfRenderingException>(() => renderer.RenderToFile(layout, path));

            Assert.Equal("既存の内容", File.ReadAllText(path));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>Rendering が解釈できない描画命令(異常系のテスト用)。</summary>
    private sealed record UnknownCommand : DrawCommand;
}
