using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Utsushi.Core.Exceptions
{
    /// <summary>
    /// 入力ファイルが有効な .xlsx として読めない(非xlsx形式・破損・パスワード保護)、
    /// またはファイル自体を開けない(存在しない・アクセス権限がない)。要件6.1, 6.2, 6.5。
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

    /// <summary>入力ファイルが読めない理由の種別(要件6.1, 6.2, 6.5)。</summary>
    public enum InvalidExcelFileReason
    {
        /// <summary>分類できないその他の読み取り失敗、またはファイル自体を開けない(存在しない・アクセス不可)。</summary>
        Unknown = 0,

        /// <summary>そもそも .xlsx(OOXML/ZIP)ではない。</summary>
        NotOpenXmlFormat,

        /// <summary>ZIP/OOXML構造が壊れている。</summary>
        Corrupted,

        /// <summary>パスワードで暗号化されている。</summary>
        PasswordProtected,

        /// <summary>ワークシートを1つも含まない。</summary>
        NoWorksheet,

        /// <summary>
        /// ファイルサイズ・展開後サイズ・シート内のセル数・共有文字列数のいずれかが
        /// 安全な処理を続けられる上限を超えている(要件6, 7)。信頼できない入力
        /// (帳票定義への登録前のExcelファイル)に対する安全弁であり、`unsupportedElements`の
        /// 設定によらず常に例外化する(パスワード保護・破損ファイルと同様、要素単位の
        /// スキップでは対処できない、ファイル全体の異常として扱う)。
        /// </summary>
        TooLarge,
    }

    /// <summary>
    /// 未対応要素(excel2pdf-consolidation 要件2)があり、その扱いがエラーのため変換を中止した。
    /// 見つかった未対応要素を、件数の上限まですべて持つ(最初の1件で止めない。要件2.2)。
    /// </summary>
    public sealed class UnsupportedWorkbookElementException : UtsushiException
    {
        /// <summary>メッセージに1行ずつ載せる未対応要素の件数。それ以降は件数だけを示す(すべては <see cref="Elements"/> と <see cref="UtsushiException.ToDiagnosticString"/>)。</summary>
        internal const int MaxElementsInMessage = 20;

        public UnsupportedWorkbookElementException(
            IReadOnlyList<UnsupportedElement> elements,
            bool isTruncated,
            string? reportCode = null,
            string? sheetName = null)
            : base(BuildMessage(elements ?? throw new ArgumentNullException(nameof(elements)), isTruncated), ProcessingStage.Parsing, reportCode, sheetName)
        {
            Elements = elements;
            IsTruncated = isTruncated;
        }

        /// <summary>見つかった未対応要素(件数の上限まで)。</summary>
        public IReadOnlyList<UnsupportedElement> Elements { get; }

        /// <summary>件数の上限に達し、含めなかった未対応要素があるかどうか。</summary>
        public bool IsTruncated { get; }

        /// <inheritdoc />
        protected override void AppendDiagnosticDetails(StringBuilder builder)
        {
            foreach (var element in Elements)
            {
                builder.Append(Environment.NewLine).Append("  ").Append(element);
            }

            if (IsTruncated)
            {
                builder.Append(Environment.NewLine).Append("  (件数の上限で打ち切り)");
            }
        }

        private static string BuildMessage(IReadOnlyList<UnsupportedElement> elements, bool isTruncated)
        {
            var builder = new StringBuilder();
            builder.Append("未対応の要素が ")
                .Append(elements.Count.ToString(CultureInfo.InvariantCulture))
                .Append(isTruncated ? " 件以上" : " 件")
                .Append("あるため中止します(UnsupportedElements を Ignore にすると、これらを出力せずに変換します)。");

            for (var i = 0; i < elements.Count && i < MaxElementsInMessage; i++)
            {
                builder.Append(Environment.NewLine).Append("  ").Append(elements[i]);
            }

            if (elements.Count > MaxElementsInMessage)
            {
                builder.Append(Environment.NewLine)
                    .Append("  ほか ")
                    .Append((elements.Count - MaxElementsInMessage).ToString(CultureInfo.InvariantCulture))
                    .Append(" 件");
            }

            return builder.ToString();
        }
    }
}
