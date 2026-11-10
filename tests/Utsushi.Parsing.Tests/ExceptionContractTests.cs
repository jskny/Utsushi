using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 呼び出し元のログ処理が依存しうる、例外の公開の契約(excel2pdf-consolidation 要件1.6、タスク8.6・8.7)。
    /// </summary>
    public sealed class ExceptionContractTests
    {
        [Fact]
        public void 処理段階の整数値は帳票定義の段階を削除した後も変えない()
        {
            // 2 は削除した帳票定義の段階(ReportDefinition)。ログに整数で残った値の意味を変えないよう、欠番にしている。
            Assert.Equal(0, (int)ProcessingStage.Unknown);
            Assert.Equal(1, (int)ProcessingStage.Parsing);
            Assert.Equal(3, (int)ProcessingStage.Substitution);
            Assert.Equal(4, (int)ProcessingStage.Layout);
            Assert.Equal(5, (int)ProcessingStage.Rendering);
            Assert.False(System.Enum.IsDefined(typeof(ProcessingStage), 2));
        }

        [Fact]
        public void 診断用の文字列は文書名をdocumentの見出しで出す()
        {
            var ex = new InvalidExcelFileException("壊れています。", InvalidExcelFileReason.Corrupted, documentName: "請求書_0001");

            Assert.Equal("請求書_0001", ex.DocumentName);
            Assert.Contains("document=請求書_0001", ex.ToDiagnosticString(), System.StringComparison.Ordinal);
            Assert.DoesNotContain("reportCode=", ex.ToDiagnosticString(), System.StringComparison.Ordinal);
        }
    }
}
