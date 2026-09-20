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

        /// <summary>解決できなかったフォント名。</summary>
        public string FontName { get; }
    }
}
