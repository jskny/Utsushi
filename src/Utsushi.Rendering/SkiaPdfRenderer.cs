using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SkiaSharp;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering
{
    /// <summary>
    /// SkiaSharp の <see cref="SKDocument"/>(PDFバックエンド)を用いた <see cref="IPdfRenderer"/> の実装。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 入力は <see cref="PagedLayout"/> のみで、Excel やレイアウト計算の知識を持たない。
    /// 1ページ = 1 <see cref="SKCanvas"/> とし、命令は背景 → 罫線 → テキスト → 画像・図形の順
    /// (Layout が並べた順)に描画する。画像・図形はExcelと同様に他のセル内容より最前面になる
    /// (要件9.3, 10.3)。
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

        /// <summary>
        /// 画像のデコード後ピクセル寸法(幅・高さ)の上限。宣言サイズが極端に大きい画像を
        /// 実際にデコードする前に拒否し、メモリを大量消費させる攻撃(いわゆるピクセル爆弾)を防ぐ
        /// (security-reviewer指摘)。自社ロゴ用途でこの上限に達することは想定していない。
        /// </summary>
        private const int MaxDecodedImageDimensionPx = 4096;

        /// <summary>
        /// <see cref="ConnectorModel.Outline"/>が<c>null</c>の接続線に補う既定の枠線
        /// (Excel上は黒い実線1ptで表示されるため。design.md参照)。
        /// </summary>
        private static readonly ShapeOutline DefaultConnectorOutline = new(ArgbColor.Black, 1.0);

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
                    DrawPage(canvas, page, layout.ReportCode, layout.SheetName);
                }
                finally
                {
                    document.EndPage();
                }
            }

            document.Close();
            skStream.Flush();
        }

        private void DrawPage(SKCanvas canvas, PageLayout page, string reportCode, string sheetName)
        {
            foreach (var command in page.Commands)
            {
                DrawSingleCommand(canvas, command, reportCode, sheetName);
            }
        }

        /// <summary>
        /// 1つの描画命令を振り分けて描画する。<see cref="DrawPage"/>の<c>foreach</c>と、
        /// <see cref="DrawGroup"/>によるグループ内子コマンドの再帰的な展開の両方から呼ばれる
        /// (design.md「Rendering レイヤー」参照)。抽象レコード型<see cref="DrawCommand"/>と
        /// 紛らわしくなるため、メソッド名は型名とは別の<c>DrawSingleCommand</c>とする。
        /// </summary>
        private void DrawSingleCommand(SKCanvas canvas, DrawCommand command, string reportCode, string sheetName)
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
                    DrawText(canvas, text, reportCode, sheetName);
                    break;

                case ImageCommand image:
                    DrawImage(canvas, image, reportCode, sheetName);
                    break;

                case ShapeCommand shape:
                    DrawShape(canvas, shape, reportCode, sheetName);
                    break;

                case ConnectorCommand connector:
                    DrawConnector(canvas, connector);
                    break;

                case GroupCommand group:
                    DrawGroup(canvas, group, reportCode, sheetName);
                    break;

                default:
                    throw new PdfRenderingException(
                        $"未知の描画命令です: {command.GetType().Name}", reportCode, sheetName);
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

        /// <summary>画像を描画する(要件9)。他のセル内容より最前面に描画される。</summary>
        private static void DrawImage(SKCanvas canvas, ImageCommand image, string reportCode, string sheetName)
        {
            // 実際にデコードする前に宣言サイズを確認し、極端に大きい画像
            // (いわゆるピクセル爆弾。数百バイトのファイルが数千万〜数億ピクセル相当を
            // 宣言することでメモリを大量消費させる攻撃)を拒否する(security-reviewer指摘)。
            var bounds = SKBitmap.DecodeBounds(image.Data);
            if (bounds.IsEmpty)
            {
                throw new PdfRenderingException(
                    $"画像(ContentType: {image.ContentType})をデコードできませんでした。", reportCode, sheetName);
            }

            if (bounds.Width > MaxDecodedImageDimensionPx || bounds.Height > MaxDecodedImageDimensionPx)
            {
                throw new PdfRenderingException(
                    $"画像(ContentType: {image.ContentType})の寸法({bounds.Width}x{bounds.Height}px)が"
                    + $"上限({MaxDecodedImageDimensionPx}px)を超えています。",
                    reportCode, sheetName);
            }

            using var bitmap = SKBitmap.Decode(image.Data);
            if (bitmap is null)
            {
                throw new PdfRenderingException(
                    $"画像(ContentType: {image.ContentType})をデコードできませんでした。", reportCode, sheetName);
            }

            var skRect = ToSkRect(image.Rect);
            var hasRotation = Math.Abs(image.RotationDegrees) > double.Epsilon;

            if (hasRotation)
            {
                canvas.Save();
                var centerX = (skRect.Left + skRect.Right) / 2f;
                var centerY = (skRect.Top + skRect.Bottom) / 2f;
                canvas.RotateDegrees((float)image.RotationDegrees, centerX, centerY);
            }

            try
            {
                canvas.DrawBitmap(bitmap, skRect);
            }
            finally
            {
                if (hasRotation)
                {
                    canvas.Restore();
                }
            }
        }

        /// <summary>図形を描画する(要件10)。他のセル内容より最前面に描画される。</summary>
        private void DrawShape(SKCanvas canvas, ShapeCommand shape, string reportCode, string sheetName)
        {
            var skRect = ToSkRect(shape.Rect);
            var hasRotation = Math.Abs(shape.RotationDegrees) > double.Epsilon;

            if (hasRotation)
            {
                canvas.Save();
                var centerX = (skRect.Left + skRect.Right) / 2f;
                var centerY = (skRect.Top + skRect.Bottom) / 2f;
                canvas.RotateDegrees((float)shape.RotationDegrees, centerX, centerY);
            }

            try
            {
                if (shape.Fill is { } fill)
                {
                    using var fillPath = ShapeGeometryBuilder.Build(shape.Preset, shape.AdjustmentValues, skRect);
                    using var fillPaint = CreateShapeFillPaint(fill, skRect);
                    canvas.DrawPath(fillPath, fillPaint);
                }

                if (shape.Outline is { } outline)
                {
                    // callout1/2/3は本体と塗りつぶしを持たない引き出し線が別ジオメトリになるため、
                    // 塗りつぶし用(Build)とは別に枠線用のジオメトリを組み立てる(それ以外の
                    // プリセットはBuildと同じ形状を返す)。
                    using var outlinePath = ShapeGeometryBuilder.BuildOutline(shape.Preset, shape.AdjustmentValues, skRect);
                    using var outlinePaint = new SKPaint
                    {
                        Color = ToSkColor(outline.Color),
                        Style = SKPaintStyle.Stroke,
                        StrokeWidth = (float)outline.WidthPt,
                        IsAntialias = true,
                    };
                    canvas.DrawPath(outlinePath, outlinePaint);
                }

                // テキストは図形本体と同じ回転変換の内側で描画することで、回転が正しく反映される
                // (design.md「Rendering レイヤー」参照)。ShapeTextLineはクリップ矩形を持たないため
                // TextCommandへの変換ではClipRectをnullにする。
                foreach (var line in shape.TextLines)
                {
                    DrawText(
                        canvas,
                        new TextCommand(line.Origin, line.Text, line.Font, line.Anchor, ClipRect: null),
                        reportCode,
                        sheetName);
                }
            }
            finally
            {
                if (hasRotation)
                {
                    canvas.Restore();
                }
            }
        }

        /// <summary>接続線を描画する(要件10.9)。塗りつぶし・テキストを持たない。</summary>
        private static void DrawConnector(SKCanvas canvas, ConnectorCommand connector)
        {
            var skRect = ToSkRect(connector.Rect);
            var isResolved = connector.ResolvedStart is not null && connector.ResolvedEnd is not null;

            // 接続点解決(要件10.11)で両端点が絶対座標として確定している場合、その座標が
            // 最終的な見た目そのものであり、Rect中心を軸にした追加の回転はかえって位置を
            // ずらしてしまう(design.md「未決事項」参照)ため適用しない。
            var hasRotation = !isResolved && Math.Abs(connector.RotationDegrees) > double.Epsilon;

            if (hasRotation)
            {
                canvas.Save();
                var centerX = (skRect.Left + skRect.Right) / 2f;
                var centerY = (skRect.Top + skRect.Bottom) / 2f;
                canvas.RotateDegrees((float)connector.RotationDegrees, centerX, centerY);
            }

            try
            {
                var resolvedStart = connector.ResolvedStart is { } start ? ToSkPoint(start) : (SKPoint?)null;
                var resolvedEnd = connector.ResolvedEnd is { } end ? ToSkPoint(end) : (SKPoint?)null;
                using var path = ConnectorGeometryBuilder.Build(
                    connector.Preset, connector.FlipHorizontal, connector.FlipVertical, skRect, resolvedStart, resolvedEnd);

                // Outlineが無い接続線にも、Excel上の既定の黒い実線1ptを補う(design.md参照)。
                var outline = connector.Outline ?? DefaultConnectorOutline;
                using var paint = new SKPaint
                {
                    Color = ToSkColor(outline.Color),
                    Style = SKPaintStyle.Stroke,
                    StrokeWidth = (float)outline.WidthPt,
                    IsAntialias = true,
                };
                canvas.DrawPath(path, paint);
            }
            finally
            {
                if (hasRotation)
                {
                    canvas.Restore();
                }
            }
        }

        /// <summary>
        /// グループ化された図形を描画する(要件10.10)。グループ自身の回転を<c>canvas</c>の
        /// 座標変換として適用したうえで、子要素(すでにページ座標へ変換済み)を出現順に
        /// 再帰的に描画する。子要素個別の回転は、この変換がすでに適用された座標系の内側で
        /// さらに<c>Save</c>/<c>RotateDegrees</c>/<c>Restore</c>するため、グループの回転と
        /// 子要素個別の回転が<c>canvas</c>の変換行列のスタックにより正しく合成される
        /// (design.md「Rendering レイヤー」参照)。
        /// </summary>
        private void DrawGroup(SKCanvas canvas, GroupCommand group, string reportCode, string sheetName)
        {
            var hasRotation = Math.Abs(group.RotationDegrees) > double.Epsilon;

            if (hasRotation)
            {
                canvas.Save();
                canvas.RotateDegrees((float)group.RotationDegrees, (float)group.Center.X, (float)group.Center.Y);
            }

            try
            {
                foreach (var child in group.Children)
                {
                    DrawSingleCommand(canvas, child, reportCode, sheetName);
                }
            }
            finally
            {
                if (hasRotation)
                {
                    canvas.Restore();
                }
            }
        }

        private static SKPaint CreateShapeFillPaint(ShapeFill fill, SKRect rect)
        {
            var paint = new SKPaint
            {
                Style = SKPaintStyle.Fill,
                IsAntialias = true,
            };

            switch (fill)
            {
                case SolidShapeFill solid:
                    paint.Color = ToSkColor(solid.Color);
                    break;

                case LinearGradientShapeFill gradient:
                    var (colors, positions) = ToShaderStops(gradient.Stops);
                    var (start, end) = GradientEndpoints(rect, gradient.AngleDegrees);
                    paint.Shader = SKShader.CreateLinearGradient(start, end, colors, positions, SKShaderTileMode.Clamp);
                    break;

                case RadialGradientShapeFill radial:
                    var (radialColors, radialPositions) = ToShaderStops(radial.Stops);
                    var center = new SKPoint(
                        rect.Left + (rect.Width * (float)radial.CenterFraction.X),
                        rect.Top + (rect.Height * (float)radial.CenterFraction.Y));
                    var radius = (float)(Math.Sqrt((rect.Width * rect.Width) + (rect.Height * rect.Height)) / 2.0);
                    paint.Shader = SKShader.CreateRadialGradient(center, radius, radialColors, radialPositions, SKShaderTileMode.Clamp);
                    break;
            }

            return paint;
        }

        /// <summary>
        /// <see cref="GradientStop"/>のリストを、位置の昇順に並べたSkiaSharpのシェーダー引数
        /// (色配列・位置配列)に変換する。<c>a:gsLst</c>の並び順はファイルの記述順であり
        /// 昇順とは限らないため、ここで並べ替える(SkiaSharpは昇順を前提とするため)。
        /// </summary>
        private static (SKColor[] Colors, float[] Positions) ToShaderStops(IReadOnlyList<GradientStop> stops)
        {
            var sorted = stops.OrderBy(s => s.Position).ToList();
            var colors = new SKColor[sorted.Count];
            var positions = new float[sorted.Count];
            for (var i = 0; i < sorted.Count; i++)
            {
                colors[i] = ToSkColor(sorted[i].Color);
                positions[i] = (float)sorted[i].Position;
            }

            return (colors, positions);
        }

        /// <summary>
        /// 矩形の中心から<paramref name="angleDegrees"/>方向に伸ばした2点を、線形グラデーションの
        /// 始点・終点とする。角度によらず矩形全体を覆うよう、対角線の半分の長さぶん伸ばす。
        /// </summary>
        private static (SKPoint Start, SKPoint End) GradientEndpoints(SKRect rect, double angleDegrees)
        {
            var centerX = (rect.Left + rect.Right) / 2.0;
            var centerY = (rect.Top + rect.Bottom) / 2.0;
            var radians = angleDegrees * Math.PI / 180.0;
            var dx = Math.Cos(radians);
            var dy = Math.Sin(radians);

            var halfDiagonal = Math.Sqrt((rect.Width * rect.Width) + (rect.Height * rect.Height)) / 2.0;

            var start = new SKPoint((float)(centerX - (dx * halfDiagonal)), (float)(centerY - (dy * halfDiagonal)));
            var end = new SKPoint((float)(centerX + (dx * halfDiagonal)), (float)(centerY + (dy * halfDiagonal)));
            return (start, end);
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

        private void DrawText(SKCanvas canvas, TextCommand text, string reportCode, string sheetName)
        {
            using var font = _fontMetrics.CreateFont(text.Font, out var synthesizeBold);
            if (_options.MissingGlyphs == MissingGlyphPolicy.Error)
            {
                EnsureGlyphsAvailable(font, text.Text, reportCode, sheetName);
            }

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
        /// 描画する各文字の字形がフォントに存在することを確かめる(要件5.5)。
        /// </summary>
        /// <remarks>
        /// 字形が無い文字は glyph 0(.notdef)に変換され、豆腐(□)または空白として描画される。
        /// 制御文字はもともと字形を描かない前提のため判定しない。
        /// </remarks>
        private static void EnsureGlyphsAvailable(SKFont font, string text, string reportCode, string sheetName)
        {
            foreach (var rune in text.EnumerateRunes())
            {
                if (Rune.IsControl(rune) || font.GetGlyph(rune.Value) != 0)
                {
                    continue;
                }

                var fontName = font.Typeface?.FamilyName ?? string.Empty;
                throw new MissingGlyphException(
                    fontName,
                    rune.Value,
                    text,
                    $"文字 '{rune}'(U+{rune.Value:X4})の字形がフォント '{fontName}' にありません"
                        + $"(描画しようとした文字列の先頭: \"{Excerpt(text)}\")。そのまま出力すると豆腐(□)や空白になるため中止しました。"
                        + "字形を持つフォントを使うか、別の文字に置き換えてください。",
                    reportCode,
                    sheetName);
            }
        }

        /// <summary>
        /// 例外メッセージに載せる文字列の抜粋。宛名・住所などの個人情報がログへ丸ごと流れないよう、先頭の数文字に留める
        /// (全文は <see cref="MissingGlyphException.Text"/> で参照できる)。
        /// </summary>
        private static string Excerpt(string text)
        {
            const int MaxTextElements = 10;
            var info = new System.Globalization.StringInfo(text);
            return info.LengthInTextElements <= MaxTextElements
                ? text
                : info.SubstringByTextElements(0, MaxTextElements) + "…";
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

        private static SKPoint ToSkPoint(PointPt point) => new((float)point.X, (float)point.Y);
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

        /// <summary>
        /// フォントに字形が無い文字を描画しようとしたときの扱い(要件5.5)。既定は
        /// <see cref="MissingGlyphPolicy.Error"/>(豆腐のままPDFを出力しない)。
        /// </summary>
        public MissingGlyphPolicy MissingGlyphs { get; init; } = MissingGlyphPolicy.Error;
    }

    /// <summary>フォントに字形が無い文字を描画しようとしたときの扱い(要件5.5)。</summary>
    public enum MissingGlyphPolicy
    {
        /// <summary><see cref="Utsushi.Core.Exceptions.MissingGlyphException"/> を送出して変換を中止する(既定)。</summary>
        Error = 0,

        /// <summary>
        /// 字形の有無を確かめずに描画する。字形の無い文字は豆腐(□)または空白になる。
        /// 代替フォントで見た目を確認するだけの開発用途を想定する。
        /// </summary>
        Render,
    }
}
