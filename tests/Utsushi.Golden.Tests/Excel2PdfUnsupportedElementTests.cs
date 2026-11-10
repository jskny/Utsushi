using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Rendering;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// <see cref="Excel2Pdf"/> の未対応要素の扱い(excel2pdf-consolidation 要件2)。グラフとコメントを含むエッジケースの注文書を使う。
    /// </summary>
    public sealed class Excel2PdfUnsupportedElementTests
    {
        private static readonly string OrderForm = Path.Combine(TestPaths.EdgeCaseSamplesRoot, "edge-order-form", "template.xlsx");
        private static readonly string Invoice = TestPaths.SampleTemplate("invoice");

        [Fact]
        public void 既定では未対応要素があると保存の時点ですべてを持つ例外にする()
        {
            using var pdf = new Excel2Pdf(OrderForm, new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback() });

            var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => pdf.ToPdfBytes());

            Assert.Contains(ex.Elements, e => e.Kind == UnsupportedElementKind.Chart && e.SheetName == "注文書" && e.Anchor is not null);
            Assert.Contains(ex.Elements, e => e.Kind == UnsupportedElementKind.LegacyDrawing);
            Assert.False(ex.IsTruncated);
            Assert.Equal("template", ex.DocumentName);
            Assert.Equal("注文書", ex.SheetName);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Contains("Ignore", ex.Message);
        }

        [Fact]
        public void 既定ではファイルへの保存も出力を残さずに例外にする()
        {
            var output = Path.Combine(Path.GetTempPath(), "utsushi-test-" + System.Guid.NewGuid().ToString("N") + ".pdf");
            using var pdf = new Excel2Pdf(OrderForm, new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback() });

            Assert.Throws<UnsupportedWorkbookElementException>(() => pdf.Save(output));
            Assert.False(File.Exists(output));
        }

        [Fact]
        public void 無視を指定すれば未対応要素を出力せずに変換できる()
        {
            using var pdf = new Excel2Pdf(
                OrderForm,
                new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback(), UnsupportedElements = UnsupportedElementPolicy.Ignore });

            Assert.True(pdf.ToPdfBytes().Length > 0);
        }

        [Fact]
        public void 未対応要素の判定は値の書き込みより先に行う()
        {
            // 制御文字を含む値のような書き込みのエラーより先に、未対応要素の例外を返す(要件2.2の補足)。
            using var pdf = new Excel2Pdf(OrderForm, new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback() });
            pdf.SetText("A3", "\u0001");

            Assert.Throws<UnsupportedWorkbookElementException>(() => pdf.ToPdfBytes());
        }

        [Fact]
        public void 未対応要素の無いテンプレートは既定のままで変換できる()
        {
            using var pdf = new Excel2Pdf(Invoice, new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback() });

            Assert.True(pdf.ToPdfBytes().Length > 0);
        }

        [Fact]
        public void 既定値はエラーである()
        {
            Assert.Equal(UnsupportedElementPolicy.Error, new Excel2PdfOptions().UnsupportedElements);
            Assert.Equal(UnsupportedElementPolicy.Error, default(UnsupportedElementPolicy));
        }
    }
}
