using System;
using System.IO;
using SkiaSharp;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering;

/// <summary>
/// SkiaSharp の <see cref="SKDocument"/>(PDFバックエンド)を用いた <see cref="IPdfRenderer"/> の実装。
/// </summary>
/// <remarks>
/// <para>
/// 入力は <see cref="PagedLayout"/> のみで、Excel やレイアウト計算の知識を持たない。
/// 1ページ = 1 <see cref="SKCanvas"/> とし、命令は背景 → 罫線 → テキストの順(Layout が並べた順)に描画する。
/// </para>
/// <para>
/// 一度メモリ上のストリームへ完全に書き出し、成功した場合のみ出力先へ転送する。
/// これにより中断時に不完全なPDFが残らない(要件5.4)。
/// </para>
/// </remarks>
public sealed class SkiaPdfRenderer : IPdfRenderer
{
    /// <summary>
    /// 太字を合成するときの輪郭の太さ(フォントサイズに対する比率)。
    /// </summary>
    /// <remarks>
    /// 一般的な擬似太字(faux bold)で使われる em の3%前後に合わせている。
    /// 大きくしすぎると字が潰れ、小さいと太字に見えない。
    /// </remarks>
    private const double BoldStrokeRatio = 0.03;

    private readonly SkiaFontMetricsProvider _fontMetrics;
    private readonly PdfRenderOptions _options;

    public SkiaPdfRenderer(SkiaFontMetricsProvider fontMetrics, PdfRenderOptions? options = null)
    {
        _fontMetrics = fontMetrics ?? throw new ArgumentNullException(nameof(fontMetrics));
        _options = options ?? PdfRenderOptions.Default;
    }

    /// <inheritdoc />
    public void Render(PagedLayout layout, Stream output)
    {
        if (layout is null)
        {
            throw new ArgumentNullException(nameof(layout));
        }

        if (output is null)
        {
            throw new ArgumentNullException(nameof(output));
        }

        if (layout.Pages.Count == 0)
        {
            throw new PdfRenderingException(
                "出力するページがありません。", layout.ReportCode, layout.SheetName);
        }

        using var buffer = new MemoryStream();
        try
        {
            RenderToStream(layout, buffer);
        }
        catch (UtsushiException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new PdfRenderingException(
                $"PDFの生成に失敗しました: {ex.Message}", layout.ReportCode, layout.SheetName, ex);
        }

        // ここまで来た時点で完全なPDFがバッファ上にある。要件5.4 のため、この時点で初めて出力先へ書く。
        buffer.Position = 0;
        buffer.CopyTo(output);
        output.Flush();
    }

    /// <summary>
    /// PDFをファイルへ出力する。失敗時に不完全なファイルを残さない(要件5.4)。
    /// </summary>
    public void RenderToFile(PagedLayout layout, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("出力パスが空です。", nameof(path));
        }

        using var buffer = new MemoryStream();
        Render(layout, buffer);

        buffer.Position = 0;
        AtomicFileWriter.Write(path, buffer, layout.ReportCode, layout.SheetName);
    }

    private void RenderToStream(PagedLayout layout, Stream target)
    {
        using var skStream = new SKManagedWStream(target, disposeManagedStream: false);

        var metadata = new SKDocumentPdfMetadata
        {
            Creator = _options.Creator,
            Producer = _options.Producer,
            Title = _options.Title ?? layout.ReportCode,
            // フォントは埋め込み(サブセット化)する。PDF/A ではなく汎用のPDF 1.4 相当を出力する。
            RasterDpi = _options.RasterDpi,
            EncodingQuality = _options.EncodingQuality,
        };

        using var document = SKDocument.CreatePdf(skStream, metadata)
            ?? throw new PdfRenderingException(
                "PDFドキュメントを生成できませんでした。", layout.ReportCode, layout.SheetName);

        foreach (var page in layout.Pages)
        {
            var canvas = document.BeginPage((float)page.WidthPt, (float)page.HeightPt);
            try
            {
                DrawPage(canvas, page);
            }
            finally
            {
                document.EndPage();
            }
        }

        document.Close();
        skStream.Flush();
    }

    private void DrawPage(SKCanvas canvas, PageLayout page)
    {
        foreach (var command in page.Commands)
        {
            switch (command)
            {
                case FillRectCommand fill:
                    DrawFill(canvas, fill);
                    break;

                case LineCommand line:
                    DrawLine(canvas, line);
                    break;

                case TextCommand text:
                    DrawText(canvas, text);
                    break;

                default:
                    throw new PdfRenderingException(
                        $"未知の描画命令です: {command.GetType().Name}");
            }
        }
    }

    private static void DrawFill(SKCanvas canvas, FillRectCommand fill)
    {
        using var paint = new SKPaint
        {
            Color = ToSkColor(fill.Color),
            Style = SKPaintStyle.Fill,
            IsAntialias = false,
        };

        canvas.DrawRect(ToSkRect(fill.Rect), paint);
    }

    private static void DrawLine(SKCanvas canvas, LineCommand line)
    {
        using var paint = new SKPaint
        {
            Color = ToSkColor(line.Color),
            Style = SKPaintStyle.Stroke,
            StrokeWidth = (float)line.WidthPt,
            IsAntialias = true,
        };

        using var dash = CreateDashEffect(line.Dash, line.WidthPt);
        if (dash is not null)
        {
            paint.PathEffect = dash;
        }

        canvas.DrawLine(
            (float)line.From.X, (float)line.From.Y,
            (float)line.To.X, (float)line.To.Y,
            paint);
    }

    private void DrawText(SKCanvas canvas, TextCommand text)
    {
        using var font = _fontMetrics.CreateFont(text.Font, out var synthesizeBold);
        using var paint = new SKPaint
        {
            Color = ToSkColor(text.Font.Color),
            IsAntialias = true,
        };

        if (synthesizeBold)
        {
            // 実フォントが太字の字形を持たない場合、輪郭を塗りと一緒に描いて太字を再現する。
            // SkiaSharp に太字を合成させると PDF が Type 3 フォントになり、
            // 文字列検索ができなくなるため、この方式で CID TrueType 埋め込みを保つ。
            paint.Style = SKPaintStyle.StrokeAndFill;
            paint.StrokeWidth = (float)(text.Font.SizePt * BoldStrokeRatio);
        }

        var width = SkiaFontMetricsProvider.MeasureText(font, text.Text);
        var x = text.Anchor switch
        {
            TextAnchor.Right => text.Origin.X - width,
            TextAnchor.Center => text.Origin.X - (width / 2.0),
            _ => text.Origin.X,
        };

        var needsClip = text.ClipRect is not null;
        if (needsClip)
        {
            canvas.Save();
            canvas.ClipRect(ToSkRect(text.ClipRect!.Value));
        }

        try
        {
            if (_options.TextRendering == PdfTextRendering.Outline)
            {
                DrawTextAsOutline(canvas, text.Text, font, paint, x, text.Origin.Y, synthesizeBold);
            }
            else
            {
                canvas.DrawText(text.Text, (float)x, (float)text.Origin.Y, font, paint);
            }

            DrawTextDecorations(canvas, text, x, width, font, paint);
        }
        finally
        {
            if (needsClip)
            {
                canvas.Restore();
            }
        }
    }

    /// <summary>
    /// 文字をグリフのアウトライン(ベクタパス)として描画する。
    /// </summary>
    /// <remarks>
    /// 位置と字形は <see cref="SKCanvas.DrawText(string, float, float, SKFont, SKPaint)"/> と同じ
    /// フォント・グリフ送り幅から求めるため、見た目は一致する。
    /// PDFにフォントを埋め込まないぶんファイルサイズが小さくなる代わりに、
    /// PDF内の文字列検索・コピーはできなくなる(<see cref="PdfTextRendering"/> を参照)。
    /// </remarks>
    private static void DrawTextAsOutline(
        SKCanvas canvas, string text, SKFont font, SKPaint paint, double x, double baselineY, bool synthesizeBold)
    {
        var glyphCount = font.CountGlyphs(text);
        if (glyphCount <= 0)
        {
            return;
        }

        var glyphs = new ushort[glyphCount];
        font.GetGlyphs(text, glyphs);

        var positions = new SKPoint[glyphCount];
        font.GetGlyphPositions(glyphs, positions, new SKPoint((float)x, (float)baselineY));

        using var combined = new SKPath();
        for (var i = 0; i < glyphCount; i++)
        {
            using var glyphPath = font.GetGlyphPath(glyphs[i]);
            if (glyphPath is null || glyphPath.IsEmpty)
            {
                // 空白などアウトラインを持たないグリフ。
                continue;
            }

            glyphPath.Transform(SKMatrix.CreateTranslation(positions[i].X, positions[i].Y));
            combined.AddPath(glyphPath);
        }

        if (combined.IsEmpty)
        {
            return;
        }

        using var fill = new SKPaint
        {
            Color = paint.Color,
            Style = synthesizeBold ? SKPaintStyle.StrokeAndFill : SKPaintStyle.Fill,
            IsAntialias = true,
            StrokeWidth = synthesizeBold ? paint.StrokeWidth : 0f,
        };

        canvas.DrawPath(combined, fill);
    }

    /// <summary>下線・取り消し線を描画する(要件4.1)。</summary>
    private static void DrawTextDecorations(
        SKCanvas canvas, TextCommand text, double x, double width, SKFont font, SKPaint paint)
    {
        var font_ = text.Font;
        if (font_.Underline == UnderlineStyle.None && !font_.Strike)
        {
            return;
        }

        var metrics = font.Metrics;
        var thickness = metrics.UnderlineThickness ?? (float)(font_.SizePt / 14.0);

        // 太字を合成している場合は、下線・取消線も同じだけ太くして字面と釣り合わせる。
        if (paint.Style == SKPaintStyle.StrokeAndFill)
        {
            thickness += paint.StrokeWidth;
        }

        using var linePaint = new SKPaint
        {
            Color = paint.Color,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = thickness,
            IsAntialias = true,
        };

        if (font_.Underline != UnderlineStyle.None)
        {
            var position = metrics.UnderlinePosition ?? (float)(font_.SizePt * 0.1);
            var y = (float)(text.Origin.Y + position);
            canvas.DrawLine((float)x, y, (float)(x + width), y, linePaint);

            if (font_.Underline is UnderlineStyle.Double or UnderlineStyle.DoubleAccounting)
            {
                var y2 = y + (thickness * 2.0f);
                canvas.DrawLine((float)x, y2, (float)(x + width), y2, linePaint);
            }
        }

        if (font_.Strike)
        {
            var strikePosition = metrics.StrikeoutPosition ?? (float)(-font_.SizePt * 0.3);
            var y = (float)(text.Origin.Y + strikePosition);
            canvas.DrawLine((float)x, y, (float)(x + width), y, linePaint);
        }
    }

    /// <summary>
    /// 破線パターンを作る。間隔は線幅に比例させ、太い罫線でも破線に見えるようにする。
    /// </summary>
    private static SKPathEffect? CreateDashEffect(LineDashStyle dash, double widthPt)
    {
        var unit = (float)Math.Max(widthPt, 0.5);

        var intervals = dash switch
        {
            LineDashStyle.Dot => new[] { unit, unit * 2f },
            LineDashStyle.Dash => new[] { unit * 4f, unit * 2f },
            LineDashStyle.DashDot => new[] { unit * 4f, unit * 2f, unit, unit * 2f },
            LineDashStyle.DashDotDot => new[] { unit * 4f, unit * 2f, unit, unit * 2f, unit, unit * 2f },
            _ => null,
        };

        return intervals is null ? null : SKPathEffect.CreateDash(intervals, 0f);
    }

    private static SKColor ToSkColor(ArgbColor color) => new(color.R, color.G, color.B, color.A);

    private static SKRect ToSkRect(RectPt rect) =>
        new((float)rect.Left, (float)rect.Top, (float)rect.Right, (float)rect.Bottom);
}

/// <summary>
/// PDF内の文字の持たせ方。
/// </summary>
/// <remarks>
/// <para>
/// SkiaSharp の PDF バックエンド(NuGet で配布されるネイティブビルド)はフォントの
/// <b>サブセット化を行わず、使用フォントを丸ごと埋め込む</b>。日本語フォントは数MBあるため、
/// 1ページの帳票でも出力PDFが数MBになる。SkiaSharp 2.88 / 3.x のいずれでも同じ挙動を確認している。
/// </para>
/// <para>
/// 帳票の配布サイズが問題になる場合は <see cref="Outline"/> を選ぶ。
/// 見た目は <see cref="EmbedFont"/> と同一だが、PDF内の文字列検索・コピー・
/// テキスト抽出ができなくなる点に注意すること。
/// </para>
/// </remarks>
public enum PdfTextRendering
{
    /// <summary>
    /// フォントを埋め込んで文字として出力する(既定)。文字列検索・コピーが可能。
    /// </summary>
    EmbedFont = 0,

    /// <summary>
    /// 文字をベクタのアウトラインとして出力する。ファイルサイズは大幅に小さくなるが、
    /// PDF内の文字は検索・コピーできなくなる。
    /// </summary>
    Outline,
}

/// <summary>PDF出力のオプション。</summary>
/// <param name="Title">PDFのタイトル。null の場合は帳票コードを使う。</param>
/// <param name="Creator">作成アプリケーション名。</param>
/// <param name="Producer">生成ライブラリ名。</param>
/// <param name="RasterDpi">ラスタ化が必要な要素の解像度。</param>
/// <param name="EncodingQuality">画像エンコード品質。</param>
/// <param name="TextRendering">文字の持たせ方。</param>
public sealed record PdfRenderOptions(
    string? Title,
    string Creator,
    string Producer,
    float RasterDpi,
    int EncodingQuality,
    PdfTextRendering TextRendering)
{
    public static PdfRenderOptions Default { get; } = new(
        Title: null,
        Creator: "Utsushi",
        Producer: "Utsushi (SkiaSharp)",
        RasterDpi: 300f,
        EncodingQuality: 100,
        TextRendering: PdfTextRendering.EmbedFont);

    /// <summary>フォントを埋め込まず、文字をアウトラインとして出力する設定。</summary>
    public static PdfRenderOptions OutlineText { get; } = Default with { TextRendering = PdfTextRendering.Outline };
}
