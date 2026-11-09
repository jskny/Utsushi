using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Rendering;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// エッジケースのサンプル(利用者が作りがちなテンプレート)で、書式の近似を診断で返すこと(excel2pdf-consolidation 要件7)。
    /// </summary>
    public sealed class EdgeCaseApproximationTests
    {
        private static string Template(string code) => Path.Combine(TestPaths.EdgeCaseSamplesRoot, code, "template.xlsx");

        private static Excel2PdfDiagnostics Check(string code)
        {
            using var pdf = new Excel2Pdf(Template(code), new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback() });
            return pdf.Check();
        }

        [Fact]
        public void 注文書の均等割り付けと選択範囲内で中央を返す()
        {
            var result = Check("edge-order-form");

            Assert.Contains(result.FormatApproximations, a => a.Kind == FormatApproximationKind.HorizontalAlignment
                && a.Location == CellRange.Parse("B24:B24") && a.Detail == "horizontal=Distributed");
            Assert.Contains(result.FormatApproximations, a => a.Kind == FormatApproximationKind.HorizontalAlignment
                && a.Location == CellRange.Parse("B25:B25") && a.Detail == "horizontal=CenterContinuous");
            Assert.Contains(result.FormatApproximations, a => a.Kind == FormatApproximationKind.TextRotation
                && a.Location == CellRange.Parse("A24:A28") && a.Detail == "textRotation=255");
            Assert.Contains(result.FormatApproximations, a => a.Kind == FormatApproximationKind.ConditionalFormatting
                && a.Location == CellRange.Parse("G10:G20"));
            Assert.True(result.HasFormatApproximations);
            Assert.All(result.FormatApproximations, a => Assert.Equal("注文書", a.SheetName));
        }

        [Fact]
        public void 注文書のリッチテキストを返し値を書き込んだセルでは返さない()
        {
            Assert.Contains(Check("edge-order-form").FormatApproximations,
                a => a.Kind == FormatApproximationKind.RichText && a.Location == CellRange.Parse("B6:D6"));

            using var pdf = new Excel2Pdf(Template("edge-order-form"), new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback() });
            pdf.SetText("B6", "書き込んだ文字列");
            Assert.DoesNotContain(pdf.Check().FormatApproximations,
                a => a.Kind == FormatApproximationKind.RichText && a.Location == CellRange.Parse("B6:D6"));
        }

        [Fact]
        public void 工程表の90度回転した日付見出しを返す()
        {
            var result = Check("edge-schedule");

            Assert.Contains(result.FormatApproximations, a => a.Kind == FormatApproximationKind.TextRotation && a.Detail == "textRotation=90");
        }
    }
}
