using System;

namespace Utsushi.Core.Exceptions
{
    /// <summary>
    /// 入力ファイルが有効な .xlsx として読めない(非xlsx形式・破損・パスワード保護)。要件6.1, 6.2。
    /// </summary>
    public sealed class InvalidExcelFileException : UtsushiException
    {
        public InvalidExcelFileException(
            string message,
            InvalidExcelFileReason reason,
            string? reportCode = null,
            Exception? innerException = null)
            : base(message, ProcessingStage.Parsing, reportCode, innerException: innerException)
        {
            Reason = reason;
        }

        /// <summary>読み込みに失敗した理由の種別。</summary>
        public InvalidExcelFileReason Reason { get; }
    }

    /// <summary>入力ファイルが読めない理由の種別(要件6.1, 6.2)。</summary>
    public enum InvalidExcelFileReason
    {
        /// <summary>分類できないその他の読み取り失敗。</summary>
        Unknown = 0,

        /// <summary>そもそも .xlsx(OOXML/ZIP)ではない。</summary>
        NotOpenXmlFormat,

        /// <summary>ZIP/OOXML構造が壊れている。</summary>
        Corrupted,

        /// <summary>パスワードで暗号化されている。</summary>
        PasswordProtected,

        /// <summary>ワークシートを1つも含まない。</summary>
        NoWorksheet,
    }

    /// <summary>
    /// 帳票定義が想定していない要素(サポート外の図形・グラフ・外部参照など)を検出した。要件1.5。
    /// </summary>
    public sealed class UnsupportedWorkbookElementException : UtsushiException
    {
        public UnsupportedWorkbookElementException(
            string message,
            string elementKind,
            string? reportCode = null,
            string? sheetName = null)
            : base(message, ProcessingStage.Parsing, reportCode, sheetName)
        {
            ElementKind = elementKind;
        }

        /// <summary>検出したサポート外要素の種別(例: "Drawing", "Chart", "ExternalReference")。</summary>
        public string ElementKind { get; }
    }
}
