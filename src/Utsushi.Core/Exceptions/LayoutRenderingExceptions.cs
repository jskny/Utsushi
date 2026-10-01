using System;

namespace Utsushi.Core.Exceptions
{
    /// <summary>
    /// レイアウト計算に失敗した(印字可能領域が確保できない、列幅が負になる等)。
    /// </summary>
    public sealed class LayoutComputationException : UtsushiException
    {
        public LayoutComputationException(
            string message,
            string? reportCode = null,
            string? sheetName = null,
            CellAddress? cellAddress = null,
            Exception? innerException = null)
            : base(message, ProcessingStage.Layout, reportCode, sheetName, cellAddress, innerException)
        {
        }
    }

    /// <summary>
    /// PDF描画に失敗した。要件5.4(不完全なPDFを残さない)に伴って送出される。
    /// </summary>
    public sealed class PdfRenderingException : UtsushiException
    {
        public PdfRenderingException(
            string message,
            string? reportCode = null,
            string? sheetName = null,
            Exception? innerException = null)
            : base(message, ProcessingStage.Rendering, reportCode, sheetName, innerException: innerException)
        {
        }
    }

    /// <summary>
    /// 帳票が要求するフォントが実行環境に存在しない。
    /// </summary>
    /// <remarks>
    /// 実行時のフォールバックによる見た目崩れを避けるため、フォールバックせず明示的に失敗させる
    /// (design.md「Rendering レイヤー」)。
    /// </remarks>
    public sealed class FontNotAvailableException : UtsushiException
    {
        public FontNotAvailableException(string fontName, string message, string? reportCode = null)
            : base(message, ProcessingStage.Rendering, reportCode)
        {
            FontName = fontName;
        }

        /// <summary>
        /// 帳票コード・シート名・処理段階を補って作る。
        /// </summary>
        /// <remarks>
        /// フォントの解決(<c>FontResolver</c>)は帳票を知らないため、帳票コード・シート名を持たない例外を投げる。
        /// ファサードがそれを捕捉し、どの帳票のどの段階(レイアウトの文字幅計測か、PDFの描画か)で起きたかを補った
        /// この例外で包み直す(要件6.4)。元の例外は <see cref="Exception.InnerException"/> に入れる。
        /// </remarks>
        public FontNotAvailableException(
            string fontName,
            string message,
            ProcessingStage stage,
            string? reportCode,
            string? sheetName,
            Exception? innerException)
            : base(message, stage, reportCode, sheetName, innerException: innerException)
        {
            FontName = fontName;
        }

        /// <summary>解決できなかったフォント名。</summary>
        public string FontName { get; }
    }

    /// <summary>
    /// 描画する文字の字形が、解決したフォントに存在しない。要件5.5。
    /// </summary>
    /// <remarks>
    /// 字形の無い文字は .notdef(豆腐または空白)として描画され、宛名などが化けたまま出力されてしまう。
    /// これを避けるため、既定では描画せずにエラーとする(<c>PdfRenderOptions.MissingGlyphs</c>)。
    /// </remarks>
    public sealed class MissingGlyphException : UtsushiException
    {
        public MissingGlyphException(
            string fontName,
            int codePoint,
            string text,
            string message,
            string? reportCode = null,
            string? sheetName = null)
            : base(message, ProcessingStage.Rendering, reportCode, sheetName)
        {
            FontName = fontName;
            CodePoint = codePoint;
            Text = text;
        }

        /// <summary>字形が見つからなかったフォント名(実際に解決された書体のファミリ名)。</summary>
        public string FontName { get; }

        /// <summary>字形が無かった文字のUnicodeコードポイント。</summary>
        public int CodePoint { get; }

        /// <summary>その文字を含んでいた描画文字列(1行分)。どのセルかを探す手がかりに使う。</summary>
        public string Text { get; }
    }
}
