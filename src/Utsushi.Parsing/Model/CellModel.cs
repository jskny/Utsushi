using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>セルに格納されている値の型。</summary>
    public enum CellValueKind
    {
        /// <summary>空セル(書式のみ)。</summary>
        Blank = 0,
        Text,
        Number,
        Boolean,
        /// <summary>Excelが保持していたエラー値(#REF! 等)。</summary>
        Error,
    }

    /// <summary>
    /// 1セル分の値と書式。
    /// </summary>
    /// <remarks>
    /// <paramref name="Value"/> は「Excelが保持している生の値」を文字列化したもの。
    /// 表示文字列(数値書式の適用結果)は <see cref="FormattedValue"/> に持つ。
    /// 数式セルは、OOXMLにキャッシュされている計算結果の値のみを読み取る(数式は評価しない)。
    /// </remarks>
    public sealed record CellModel(
        string? Value,
        CellValueKind ValueKind,
        CellStyle Style,
        string? FormattedValue = null,
        bool HasFormula = false)
    {
        /// <summary>描画に使うべき表示文字列。</summary>
        public string? DisplayValue => FormattedValue ?? Value;

        public bool IsBlank => ValueKind == CellValueKind.Blank || string.IsNullOrEmpty(DisplayValue);

        /// <summary>
        /// 数値書式の表示に使ったセクションの色の指定(<c>[Red]</c> など。要件4.12)。指定が無ければ null で、
        /// フォントの色(<see cref="CellStyle.Font"/>)で描く。
        /// </summary>
        public ArgbColor? FormatColor { get; init; }

        /// <summary>
        /// 一部またはすべての文字の書式(リッチテキストのラン)がセルの書式と異なるかどうか。Utsushi はセル全体をセルの書式で描く
        /// (書式の近似として診断で返す。excel2pdf-consolidation 要件7.2(d))。値を書き込むと消える。
        /// </summary>
        public bool HasRichText { get; init; }

        /// <summary>
        /// 数値書式を解釈できない、または書式で表せない値のため、標準の表示形式で表示したかどうか
        /// (書式の近似として診断で返す。excel2pdf-consolidation 要件7.2(e))。
        /// </summary>
        public bool NumberFormatFellBack { get; init; }

        /// <summary>
        /// 値のみを差し替えた複製を返す(書式は変更しない。要件2.2)。差し込んだ文字列は数値書式で表示しないため、
        /// 数値書式の色(<see cref="FormatColor"/>)は消す(要件4.12)。
        /// </summary>
        /// <summary>
        /// 数値に置き換えた複製を返す(書式は変更しない。要件14.3)。表示文字列と色は呼び出し元がセルの数値書式で求めて渡す。
        /// </summary>
        public CellModel WithNumber(string rawValue, string formattedValue, ArgbColor? formatColor) =>
            WithNumber(rawValue, formattedValue, formatColor, numberFormatFellBack: false);

        /// <summary>
        /// <see cref="WithNumber(string, string, ArgbColor?)"/> と同じ。<paramref name="numberFormatFellBack"/> は、数値書式で表示できずに
        /// 標準の表示にしたかどうか(<see cref="NumberFormatFellBack"/>)。
        /// </summary>
        public CellModel WithNumber(string rawValue, string formattedValue, ArgbColor? formatColor, bool numberFormatFellBack) =>
            this with
            {
                Value = rawValue,
                ValueKind = CellValueKind.Number,
                FormattedValue = formattedValue,
                HasFormula = false,
                FormatColor = formatColor,
                HasRichText = false,
                NumberFormatFellBack = numberFormatFellBack,
            };

        public CellModel WithText(string text) =>
            this with
            {
                Value = text,
                ValueKind = CellValueKind.Text,
                FormattedValue = text,
                HasFormula = false,
                FormatColor = null,
                HasRichText = false,
                NumberFormatFellBack = false,
            };
    }
}
