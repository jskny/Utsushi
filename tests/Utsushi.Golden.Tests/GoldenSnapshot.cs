using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// <see cref="PagedLayout"/> を、人が差分を読める行指向のテキストへ変換する。
    /// </summary>
    /// <remarks>
    /// <para>
    /// ゴールデンテストは PDF バイナリの完全一致ではなく、この「描画命令のスナップショット」を
    /// 比較対象とする(design.md「テスト戦略」)。理由は2つある。
    /// </para>
    /// <list type="bullet">
    ///   <item>PDFバイナリは生成日時・圧縮・SkiaSharpのバージョンで変わるため、意味のない差分が出る。</item>
    ///   <item>座標・テキスト・罫線の差分を行単位で読めるため、崩れた箇所を直接特定できる。</item>
    /// </list>
    /// <para>
    /// 座標は 0.01pt(約0.0035mm)に丸める。帳票定義の許容誤差(既定0.5mm)より十分細かく、
    /// 浮動小数の最下位ビットの揺れは吸収できる粒度である。
    /// </para>
    /// </remarks>
    internal static class GoldenSnapshot
    {
        /// <summary>座標の丸め桁数(小数以下2桁 = 0.01pt)。</summary>
        private const int CoordinatePrecision = 2;

        public static string Create(PagedLayout layout)
        {
            var sb = new StringBuilder();
            sb.Append("report=").Append(layout.ReportCode)
              .Append(" sheet=").Append(layout.SheetName)
              .Append(" pages=").Append(layout.PageCount)
              .AppendLine();

            foreach (var page in layout.Pages)
            {
                sb.AppendLine();
                sb.Append("## page ").Append(page.PageNumber)
                  .Append(" paper=").Append(page.Paper.Name)
                  .Append('/').Append(page.Orientation)
                  .Append(" size=").Append(N(page.WidthPt)).Append('x').Append(N(page.HeightPt))
                  .Append(" scale=").Append(N(page.ScaleFactor))
                  .Append(" rows=").Append(page.RowRange.First).Append('-').Append(page.RowRange.Last)
                  .Append(" cols=").Append(page.ColumnRange.First).Append('-').Append(page.ColumnRange.Last)
                  .AppendLine();

                foreach (var line in Describe(page.Commands))
                {
                    sb.AppendLine(line);
                }
            }

            return sb.ToString();
        }

        private static IEnumerable<string> Describe(IReadOnlyList<DrawCommand> commands)
        {
            foreach (var command in commands)
            {
                yield return command switch
                {
                    FillRectCommand fill =>
                        $"fill   rect={Rect(fill.Rect)} color={fill.Color}",

                    LineCommand line =>
                        $"line   from={Point(line.From)} to={Point(line.To)} "
                        + $"color={line.Color} width={N(line.WidthPt)} dash={line.Dash}",

                    TextCommand text =>
                        $"text   origin={Point(text.Origin)} anchor={text.Anchor} "
                        + $"font={Font(text)} clip={(text.ClipRect is { } c ? Rect(c) : "none")} "
                        + $"value={Quote(text.Text)}",

                    ImageCommand image =>
                        $"image  rect={Rect(image.Rect)} contentType={image.ContentType} bytes={image.Data.Length}",

                    ShapeCommand shape =>
                        $"shape  rect={Rect(shape.Rect)} preset={shape.Preset} rotation={N(shape.RotationDegrees)} "
                        + $"fill={Fill(shape.Fill)} outline={Outline(shape.Outline)} "
                        + $"text=[{string.Join(";", shape.TextLines.Select(ShapeTextLine))}]",

                    _ => $"unknown {command.GetType().Name}",
                };
            }
        }

        private static string Fill(ShapeFill? fill) => fill switch
        {
            SolidShapeFill solid => $"solid:{solid.Color}",
            LinearGradientShapeFill gradient =>
                $"linGradient:{GradientStops(gradient.Stops)}@{N(gradient.AngleDegrees)}",
            RadialGradientShapeFill radial =>
                $"radGradient:{GradientStops(radial.Stops)}@center({N(radial.CenterFraction.X)},{N(radial.CenterFraction.Y)})",
            _ => "none",
        };

        private static string GradientStops(IReadOnlyList<GradientStop> stops) =>
            string.Join(",", stops.Select(s => $"{N(s.Position)}:{s.Color}"));

        private static string Outline(ShapeOutline? outline) =>
            outline is { } o ? $"{o.Color}/{N(o.WidthPt)}pt" : "none";

        private static string ShapeTextLine(ShapeTextLine line) =>
            $"origin={Point(line.Origin)} anchor={line.Anchor} font={Font(line.Font)} value={Quote(line.Text)}";

        private static string Font(TextCommand text) => Font(text.Font);

        private static string Font(FontStyle font)
        {
            var flags = new List<string>(4);
            if (font.Bold) { flags.Add("bold"); }
            if (font.Italic) { flags.Add("italic"); }
            if (font.Underline != UnderlineStyle.None) { flags.Add("underline:" + font.Underline); }
            if (font.Strike) { flags.Add("strike"); }

            var suffix = flags.Count == 0 ? string.Empty : "," + string.Join(",", flags);
            return $"{font.Name}@{N(font.SizePt)}pt/{font.Color}{suffix}";
        }

        private static string Rect(RectPt rect) =>
            $"[{N(rect.Left)},{N(rect.Top)} {N(rect.Width)}x{N(rect.Height)}]";

        private static string Point(PointPt point) => $"({N(point.X)},{N(point.Y)})";

        private static string Quote(string value) =>
            "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n") + "\"";

        /// <summary>座標・寸法を丸めて出力する。-0 表記を避けるため 0 に正規化する。</summary>
        private static string N(double value)
        {
            var rounded = Math.Round(value, CoordinatePrecision, MidpointRounding.AwayFromZero);
            if (rounded == 0.0)
            {
                rounded = 0.0;
            }

            return rounded.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
