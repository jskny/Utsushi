using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Utsushi.Core.Exceptions
{
    /// <summary>
    /// Utsushi が送出するすべての例外の基底クラス。
    /// </summary>
    /// <remarks>
    /// 要件6.4「どの文書・どのセル・どの処理段階で失敗したかを特定できる情報を含める」に対応するため、
    /// 文書名・シート名・セル番地・処理段階を構造化プロパティとして保持する。
    /// ログ出力時は <see cref="ToDiagnosticString"/> を用いるとこれらをまとめて出力できる。
    /// </remarks>
    public abstract class UtsushiException : Exception
    {
        protected UtsushiException(
            string message,
            ProcessingStage stage,
            string? documentName = null,
            string? sheetName = null,
            CellAddress? cellAddress = null,
            Exception? innerException = null)
            : base(message, innerException)
        {
            Stage = stage;
            DocumentName = documentName;
            SheetName = sheetName;
            CellAddress = cellAddress;
        }

        /// <summary>失敗した処理段階。</summary>
        public ProcessingStage Stage { get; }

        /// <summary>関係する文書名(判明している場合)。</summary>
        public string? DocumentName { get; }

        /// <summary>関係するシート名(判明している場合)。</summary>
        public string? SheetName { get; }

        /// <summary>関係するセル番地(判明している場合)。</summary>
        public CellAddress? CellAddress { get; }

        /// <summary>ログ出力用に、保持する構造化情報をまとめた文字列を返す。</summary>
        public string ToDiagnosticString()
        {
            var parts = new List<string>(5)
        {
            "stage=" + Stage.ToString(),
        };

            if (!string.IsNullOrEmpty(DocumentName))
            {
                parts.Add("document=" + DocumentName);
            }

            if (!string.IsNullOrEmpty(SheetName))
            {
                parts.Add("sheet=" + SheetName);
            }

            if (CellAddress.HasValue)
            {
                parts.Add("cell=" + CellAddress.Value);
            }

            var sb = new StringBuilder();
            sb.Append(GetType().Name)
              .Append(": ")
              .Append(Message)
              .Append(" (")
              .Append(string.Join(", ", parts))
              .Append(')');
            AppendDiagnosticDetails(sb);
            return sb.ToString();
        }

        /// <summary>
        /// 派生クラスが <see cref="ToDiagnosticString"/> に詳細(未対応要素の一覧など)を加えるためのフック。
        /// 既定では何もしない。加える場合は、改行で始めて1行ずつ書く。
        /// </summary>
        protected virtual void AppendDiagnosticDetails(StringBuilder builder)
        {
        }

        public override string ToString() =>
            ToDiagnosticString() + Environment.NewLine + (StackTrace ?? string.Empty)
            + (InnerException is null
                ? string.Empty
                : string.Format(CultureInfo.InvariantCulture, "{0} ---> {1}", Environment.NewLine, InnerException));
    }
}
