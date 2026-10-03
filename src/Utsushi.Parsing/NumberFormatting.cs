using Utsushi.Core;
using Utsushi.Parsing.OpenXml;

namespace Utsushi.Parsing
{
    /// <summary>
    /// セルの数値書式(<c>numFmt</c>)を数値に適用する(要件4.8, 4.12)。Excel ファイルを読むときと同じ規則で、
    /// 呼び出し元が設定した数値(要件14.3)の表示文字列と色を求めるために公開する。
    /// </summary>
    public static class NumberFormatting
    {
        /// <summary>数値に書式を適用した表示文字列を返す。解釈できない書式は標準(General)の表示にする。</summary>
        /// <param name="value">数値。</param>
        /// <param name="formatCode">数値書式(<see cref="Model.CellStyle.NumberFormat"/>)。null は標準。</param>
        public static string Format(double value, string? formatCode) => NumberFormatter.FormatNumber(value, formatCode);

        /// <summary>数値の表示に使うセクションの色の指定(<c>[Red]</c> など)を返す。無ければ null。</summary>
        /// <param name="value">数値。</param>
        /// <param name="formatCode">数値書式。</param>
        public static ArgbColor? ResolveColor(double value, string? formatCode) => NumberFormatter.ResolveColor(value, formatCode);
    }
}
