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

        /// <summary>
        /// 数値に書式を適用した表示文字列・色と、書式を解釈できない(または書式で表せない値の)ために標準の表示にしたかどうかを返す
        /// (excel2pdf-consolidation 要件7.2(e))。<see cref="Format"/>・<see cref="ResolveColor"/> と同じ規則。
        /// </summary>
        /// <param name="value">数値。</param>
        /// <param name="formatCode">数値書式。</param>
        public static NumberFormatResult FormatWithFallback(double value, string? formatCode)
        {
            var text = NumberFormatter.FormatNumber(value, formatCode, out NumberFormatFallback fallback);
            return new NumberFormatResult(text, NumberFormatter.ResolveColor(value, formatCode), fallback);
        }
    }

    /// <summary>数値に書式を適用した結果。</summary>
    /// <param name="Text">表示文字列。</param>
    /// <param name="Color">色の指定。無ければ null。</param>
    /// <param name="Fallback">書式を解釈できない、または書式で表せない値のため、標準の表示にしたかどうかと、その理由。</param>
    public sealed record NumberFormatResult(string Text, ArgbColor? Color, NumberFormatFallback Fallback)
    {
        /// <summary>標準の表示にしたかどうか。</summary>
        public bool FellBackToGeneral => Fallback != NumberFormatFallback.None;
    }

    /// <summary>数値書式で表示できずに標準の表示にした理由(excel2pdf-consolidation 要件7.2(e))。</summary>
    public enum NumberFormatFallback
    {
        /// <summary>書式どおりに表示した。</summary>
        None = 0,

        /// <summary>書式を解釈できない(分数・指数など)。テンプレートの書式が原因。</summary>
        UnsupportedFormat = 1,

        /// <summary>書式で表せない値(負の経過時間・明治より前の日付の和暦など)。値が原因。</summary>
        UnrepresentableValue = 2,
    }
}
