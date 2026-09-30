using System;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// Excel固有の寸法単位を、Utsushi内部の単位(ポイント)へ換算する。
    /// </summary>
    /// <remarks>
    /// <para>
    /// Excel の列幅は「標準フォントで数字を何文字分表示できるか」という単位で保存されており、
    /// 標準フォントの最大数字幅(MDW: Maximum Digit Width、96dpiのピクセル)に依存する。
    /// ECMA-376 Part 1, 18.3.1.13 が定める換算式は次の通り:
    /// </para>
    /// <code>
    /// pixels = Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW)
    /// </code>
    /// <para>
    /// Truncate による切り捨てが2回入るため、素朴な <c>width * MDW</c> とは最大1px程度ずれる。
    /// この誤差は複数列にわたって累積し改ページ位置に影響しうるため、式をそのまま実装する。
    /// </para>
    /// <para>行高(<c>ht</c>)は元々ポイント単位で保存されているため換算は不要。</para>
    /// </remarks>
    public static class ExcelUnitConverter
    {
        /// <summary>
        /// 列幅(文字数単位)を96dpiのピクセル数へ換算する。
        /// </summary>
        /// <param name="width">OOXML の <c>col/@width</c>。</param>
        /// <param name="maxDigitWidthPx">標準フォントの最大数字幅(ピクセル)。</param>
        public static double ColumnWidthToPixels(double width, double maxDigitWidthPx)
        {
            if (maxDigitWidthPx <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(maxDigitWidthPx), maxDigitWidthPx, "最大数字幅は正の数である必要があります。");
            }

            if (width <= 0)
            {
                return 0.0;
            }

            var padding = Math.Truncate(128.0 / maxDigitWidthPx);
            return Math.Truncate(((256.0 * width) + padding) / 256.0 * maxDigitWidthPx);
        }

        /// <summary>列幅(文字数単位)をポイントへ換算する。</summary>
        public static double ColumnWidthToPoints(double width, double maxDigitWidthPx) =>
            Units.PixelsToPoints(ColumnWidthToPixels(width, maxDigitWidthPx));

        /// <summary>
        /// シートの列幅をポイントへ換算する。幅の指定が無い列(<see cref="SheetModel.GetColumnWidth"/> が NaN)は、
        /// <see cref="DefaultColumnWidthToPixels"/> で求めた既定列幅を使う。
        /// </summary>
        public static double SheetColumnWidthToPoints(SheetModel sheet, int column, double maxDigitWidthPx)
        {
            var width = sheet.GetColumnWidth(column);
            return double.IsNaN(width)
                ? Units.PixelsToPoints(DefaultColumnWidthToPixels(sheet.BaseColumnWidth, maxDigitWidthPx))
                : ColumnWidthToPoints(width, maxDigitWidthPx);
        }

        /// <summary>
        /// <c>defaultColWidth</c> が無いシートの既定列幅(96dpiのピクセル)。Excel は
        /// 「<c>baseColWidth</c> 文字分 + 余白5px」を8ピクセルの倍数に切り上げる(Calibri 11 = 最大数字幅7 で 64px =「8.43」、
        /// ＭＳ Ｐゴシック/游ゴシック 11 = 最大数字幅8 で 72px =「8.38」)。
        /// </summary>
        public static double DefaultColumnWidthToPixels(int baseColumnWidth, double maxDigitWidthPx)
        {
            if (maxDigitWidthPx <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(maxDigitWidthPx), maxDigitWidthPx, "最大数字幅は正の数である必要があります。");
            }

            var raw = (Math.Max(0, baseColumnWidth) * maxDigitWidthPx) + 5.0;
            return Math.Ceiling(raw / 8.0) * 8.0;
        }

        /// <summary>
        /// インデント1段分の幅(ポイント)。Excel はインデント1段を標準フォント3文字分として扱う。
        /// </summary>
        public static double IndentWidthToPoints(int indentLevel, double maxDigitWidthPx)
        {
            if (indentLevel <= 0)
            {
                return 0.0;
            }

            return Units.PixelsToPoints(indentLevel * 3.0 * maxDigitWidthPx);
        }

        /// <summary>
        /// セル内容の左右パディング(ポイント)。Excel は左右それぞれ2ピクセル分の余白を取る。
        /// </summary>
        public static double CellPaddingPoints => Units.PixelsToPoints(2.0);
    }
}
