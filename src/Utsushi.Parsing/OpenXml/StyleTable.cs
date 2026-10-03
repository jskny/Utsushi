using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// ワークブックのスタイル情報(<c>styles.xml</c>)を解決済みの <see cref="CellStyle"/> として引けるようにする。
    /// </summary>
    internal sealed class StyleTable
    {
        /// <summary>
        /// フォント/塗りつぶし/罫線/数値書式/セル書式(<c>cellXfs</c>)それぞれの要素数の上限。
        /// 信頼できない入力による計算量・メモリの増大を防ぐ安全弁(要件6, 7。security-reviewer指摘)。
        /// 上限を超えた分は読み取らず、それを参照する<c>styleIndex</c>は既存の「範囲外索引は
        /// 既定書式」という挙動(<see cref="GetCellStyle"/>)にそのまま従う(画像・図形の個数上限が
        /// 超過分を無視するのと同様の扱いであり、例外化はしない)。
        /// </summary>
        private const int MaxStyleTableEntries = 10_000;

        private readonly IReadOnlyList<CellStyle> _cellStyles;
        private readonly CellStyle _fallback;

        private StyleTable(IReadOnlyList<CellStyle> cellStyles, CellStyle fallback, FontStyle defaultFont)
        {
            _cellStyles = cellStyles;
            _fallback = fallback;
            DefaultFont = defaultFont;
        }

        /// <summary>ブックの標準の書式(<c>cellXfs</c> 索引0)。セルが無い位置に使う(要件1.10)。</summary>
        public CellStyle DefaultCellStyle => _fallback;

        /// <summary>ブックの標準フォント(<c>cellXfs</c> 索引0のフォント)。列幅換算の基準に用いる。</summary>
        public FontStyle DefaultFont { get; }

        public static StyleTable Create(WorkbookPart workbookPart, ColorResolver colors) =>
            Create(workbookPart, colors, MaxStyleTableEntries);

        /// <summary>
        /// <see cref="Create(WorkbookPart, ColorResolver)"/>の本体。実際の上限
        /// (<see cref="MaxStyleTableEntries"/>、既定1万件)は現実的なユニットテストでは
        /// 大量のスタイル要素を用意しないと到達できないため、<paramref name="maxEntries"/>を
        /// 明示的に指定できる形にしてテスト可能にしている。
        /// </summary>
        internal static StyleTable Create(WorkbookPart workbookPart, ColorResolver colors, int maxEntries)
        {
            var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
            if (stylesheet is null)
            {
                return new StyleTable(Array.Empty<CellStyle>(), CellStyle.Default, FontStyle.Default);
            }

            var numberFormats = ReadCustomNumberFormats(stylesheet, maxEntries);
            var fonts = ReadFonts(stylesheet, colors, maxEntries);
            var fills = ReadFills(stylesheet, colors, maxEntries);
            var borders = ReadBorders(stylesheet, colors, maxEntries);

            var resolved = new List<CellStyle>();
            foreach (var xf in stylesheet.CellFormats?.Elements<X.CellFormat>() ?? Array.Empty<X.CellFormat>())
            {
                if (resolved.Count >= maxEntries)
                {
                    break;
                }

                resolved.Add(ResolveCellFormat(xf, fonts, fills, borders, numberFormats));
            }

            var defaultFont = resolved.Count > 0 ? resolved[0].Font : (fonts.Count > 0 ? fonts[0] : FontStyle.Default);
            var fallback = resolved.Count > 0 ? resolved[0] : CellStyle.Default;
            return new StyleTable(resolved, fallback, defaultFont);
        }

        /// <summary><c>cellXfs</c> の索引から書式を取得する。範囲外の索引には既定書式を返す。</summary>
        public CellStyle GetCellStyle(int? styleIndex)
        {
            if (styleIndex is null)
            {
                return _fallback;
            }

            var index = styleIndex.Value;
            return index >= 0 && index < _cellStyles.Count ? _cellStyles[index] : _fallback;
        }

        private static CellStyle ResolveCellFormat(
            X.CellFormat xf,
            IReadOnlyList<FontStyle> fonts,
            IReadOnlyList<ArgbColor> fills,
            IReadOnlyList<BorderSet> borders,
            IReadOnlyDictionary<int, string> customNumberFormats)
        {
            var font = ResolveIndexed(fonts, xf.FontId?.Value, FontStyle.Default);
            var fill = ResolveIndexed(fills, xf.FillId?.Value, ArgbColor.Transparent);
            var border = ResolveIndexed(borders, xf.BorderId?.Value, BorderSet.None);

            string? numberFormat = null;
            if (xf.NumberFormatId?.Value is { } numFmtId)
            {
                numberFormat = customNumberFormats.TryGetValue((int)numFmtId, out var custom)
                    ? custom
                    : NumberFormatter.GetBuiltInFormat((int)numFmtId);
            }

            var alignment = xf.Alignment;
            var hAlign = MapHorizontal(alignment?.Horizontal);
            var vAlign = MapVertical(alignment?.Vertical);
            var wrap = alignment?.WrapText?.Value ?? false;
            var shrink = alignment?.ShrinkToFit?.Value ?? false;
            var indent = (int)(alignment?.Indent?.Value ?? 0U);

            return new CellStyle(font, border, hAlign, vAlign, numberFormat, fill, wrap, shrink, indent);
        }

        private static T ResolveIndexed<T>(IReadOnlyList<T> list, uint? index, T fallback) =>
            index is { } i && i < list.Count ? list[(int)i] : fallback;

        /// <summary>数値書式コード(<c>numFmt/@formatCode</c>)の長さの上限(Excel 自体の上限。要件6.8)。</summary>
        internal const int MaxNumberFormatCodeLength = 255;

        private static Dictionary<int, string> ReadCustomNumberFormats(X.Stylesheet stylesheet, int maxEntries)
        {
            var result = new Dictionary<int, string>();
            foreach (var numFmt in stylesheet.NumberingFormats?.Elements<X.NumberingFormat>() ?? Array.Empty<X.NumberingFormat>())
            {
                if (result.Count >= maxEntries)
                {
                    break;
                }

                // Excel 自体の上限(255文字)を超える書式コードは読み取らず、既定書式(標準)にフォールバックする
                // (書式はセルごとに解析するため、極端に長い書式を多数のセルに適用させて処理を長引かせるのを防ぐ。
                // 要件6.8。security-reviewer指摘)。
                if (numFmt.NumberFormatId?.Value is { } id && numFmt.FormatCode?.Value is { } code
                    && code.Length <= MaxNumberFormatCodeLength)
                {
                    result[(int)id] = code;
                }
            }

            return result;
        }

        private static List<FontStyle> ReadFonts(X.Stylesheet stylesheet, ColorResolver colors, int maxEntries)
        {
            var result = new List<FontStyle>();
            foreach (var font in stylesheet.Fonts?.Elements<X.Font>() ?? Array.Empty<X.Font>())
            {
                if (result.Count >= maxEntries)
                {
                    break;
                }

                var name = font.FontName?.Val?.Value ?? FontStyle.Default.Name;
                // NaN・無限大・範囲外のサイズは既定のサイズに戻す(要件6.9)。
                var size = OpenXmlWorkbookReader.ValidFontSizeOrDefault(font.FontSize?.Val?.Value);
                var bold = IsOn(font.Bold);
                var italic = IsOn(font.Italic);
                var strike = IsOn(font.Strike);
                var underline = MapUnderline(font.Underline);
                var color = colors.Resolve(font.Color, ArgbColor.Black);
                result.Add(new FontStyle(name, size, bold, italic, underline, strike, color));
            }

            return result;
        }

        /// <summary>
        /// <c>b</c>/<c>i</c>/<c>strike</c> は要素の存在だけで true を意味する(val属性は省略可)。
        /// </summary>
        private static bool IsOn(X.BooleanPropertyType? property) => property is not null && (property.Val?.Value ?? true);

        private static UnderlineStyle MapUnderline(X.Underline? underline)
        {
            if (underline is null)
            {
                return UnderlineStyle.None;
            }

            if (underline.Val is null)
            {
                // <u/> は単線を意味する。
                return UnderlineStyle.Single;
            }

            var value = underline.Val.Value;
            if (value == X.UnderlineValues.None) { return UnderlineStyle.None; }
            if (value == X.UnderlineValues.Double) { return UnderlineStyle.Double; }
            if (value == X.UnderlineValues.SingleAccounting) { return UnderlineStyle.SingleAccounting; }
            if (value == X.UnderlineValues.DoubleAccounting) { return UnderlineStyle.DoubleAccounting; }
            return UnderlineStyle.Single;
        }

        private static List<ArgbColor> ReadFills(X.Stylesheet stylesheet, ColorResolver colors, int maxEntries)
        {
            var result = new List<ArgbColor>();
            foreach (var fill in stylesheet.Fills?.Elements<X.Fill>() ?? Array.Empty<X.Fill>())
            {
                if (result.Count >= maxEntries)
                {
                    break;
                }

                var pattern = fill.PatternFill;
                if (pattern?.PatternType is null || pattern.PatternType.Value == X.PatternValues.None)
                {
                    result.Add(ArgbColor.Transparent);
                    continue;
                }

                if (pattern.PatternType.Value == X.PatternValues.Solid)
                {
                    // 塗りつぶし(単色)では前景色が塗り色になる。
                    result.Add(colors.Resolve(pattern.ForegroundColor, ArgbColor.Transparent));
                    continue;
                }

                // 網掛けパターンは対象帳票では使用しない想定のため、背景色で近似する。
                result.Add(colors.Resolve(pattern.BackgroundColor, ArgbColor.Transparent));
            }

            return result;
        }

        private static List<BorderSet> ReadBorders(X.Stylesheet stylesheet, ColorResolver colors, int maxEntries)
        {
            var result = new List<BorderSet>();
            foreach (var border in stylesheet.Borders?.Elements<X.Border>() ?? Array.Empty<X.Border>())
            {
                if (result.Count >= maxEntries)
                {
                    break;
                }

                var diagonalDown = border.DiagonalDown?.Value ?? false;
                var diagonalUp = border.DiagonalUp?.Value ?? false;
                var diagonal = ReadEdge(border.DiagonalBorder, colors);

                result.Add(new BorderSet(
                    ReadEdge(border.LeftBorder, colors),
                    ReadEdge(border.RightBorder, colors),
                    ReadEdge(border.TopBorder, colors),
                    ReadEdge(border.BottomBorder, colors),
                    diagonalDown ? diagonal : BorderEdge.None,
                    diagonalUp ? diagonal : BorderEdge.None));
            }

            return result;
        }

        private static BorderEdge ReadEdge(X.BorderPropertiesType? edge, ColorResolver colors)
        {
            if (edge?.Style is null)
            {
                return BorderEdge.None;
            }

            var style = MapBorderStyle(edge.Style.Value);
            if (style == BorderLineStyle.None)
            {
                return BorderEdge.None;
            }

            return new BorderEdge(style, colors.Resolve(edge.Color, ArgbColor.Black));
        }

        private static BorderLineStyle MapBorderStyle(X.BorderStyleValues value)
        {
            if (value == X.BorderStyleValues.Thin) { return BorderLineStyle.Thin; }
            if (value == X.BorderStyleValues.Medium) { return BorderLineStyle.Medium; }
            if (value == X.BorderStyleValues.Thick) { return BorderLineStyle.Thick; }
            if (value == X.BorderStyleValues.Hair) { return BorderLineStyle.Hair; }
            if (value == X.BorderStyleValues.Dotted) { return BorderLineStyle.Dotted; }
            if (value == X.BorderStyleValues.Dashed) { return BorderLineStyle.Dashed; }
            if (value == X.BorderStyleValues.DashDot) { return BorderLineStyle.DashDot; }
            if (value == X.BorderStyleValues.DashDotDot) { return BorderLineStyle.DashDotDot; }
            if (value == X.BorderStyleValues.Double) { return BorderLineStyle.Double; }
            if (value == X.BorderStyleValues.MediumDashed) { return BorderLineStyle.MediumDashed; }
            if (value == X.BorderStyleValues.MediumDashDot) { return BorderLineStyle.MediumDashDot; }
            if (value == X.BorderStyleValues.MediumDashDotDot) { return BorderLineStyle.MediumDashDotDot; }
            if (value == X.BorderStyleValues.SlantDashDot) { return BorderLineStyle.SlantDashDot; }
            return BorderLineStyle.None;
        }

        private static HorizontalAlignment MapHorizontal(EnumValue<X.HorizontalAlignmentValues>? value)
        {
            if (value is null)
            {
                return HorizontalAlignment.General;
            }

            var v = value.Value;
            if (v == X.HorizontalAlignmentValues.Left) { return HorizontalAlignment.Left; }
            if (v == X.HorizontalAlignmentValues.Center) { return HorizontalAlignment.Center; }
            if (v == X.HorizontalAlignmentValues.Right) { return HorizontalAlignment.Right; }
            if (v == X.HorizontalAlignmentValues.Fill) { return HorizontalAlignment.Fill; }
            if (v == X.HorizontalAlignmentValues.Justify) { return HorizontalAlignment.Justify; }
            if (v == X.HorizontalAlignmentValues.CenterContinuous) { return HorizontalAlignment.CenterContinuous; }
            if (v == X.HorizontalAlignmentValues.Distributed) { return HorizontalAlignment.Distributed; }
            return HorizontalAlignment.General;
        }

        private static VerticalAlignment MapVertical(EnumValue<X.VerticalAlignmentValues>? value)
        {
            if (value is null)
            {
                // Excel の既定は下揃え。
                return VerticalAlignment.Bottom;
            }

            var v = value.Value;
            if (v == X.VerticalAlignmentValues.Top) { return VerticalAlignment.Top; }
            if (v == X.VerticalAlignmentValues.Center) { return VerticalAlignment.Center; }
            if (v == X.VerticalAlignmentValues.Justify) { return VerticalAlignment.Justify; }
            if (v == X.VerticalAlignmentValues.Distributed) { return VerticalAlignment.Distributed; }
            return VerticalAlignment.Bottom;
        }
    }
}
