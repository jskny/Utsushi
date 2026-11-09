using System.IO;
using System.Linq;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Utsushi.TestSupport;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 書式の近似の判断に必要な書式の読み取り(excel2pdf-consolidation design.md 3.5「検知の分担」)。
    /// </summary>
    public sealed class FormatApproximationReadingTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        [Fact]
        public void 図形の文字の縦書きと回転を読み取る()
        {
            var body = ShapeWorkbookFixtures.TextBody("縦");
            body.BodyProperties!.Vertical = A.TextVerticalValues.EastAsianVetical;
            body.BodyProperties.Rotation = 5400000;
            var path = ShapeWorkbookFixtures.CreateWorkbook(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, textBody: body));
            try
            {
                var shape = Assert.Single(_reader.ReadFile(path).Sheets[0].DrawingObjects.OfType<ShapeModel>());

                Assert.NotNull(shape.Text!.Vertical);
                Assert.Equal(90.0, shape.Text.RotationDegrees, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 横書きの図形の文字は縦書きとして読まない()
        {
            var body = ShapeWorkbookFixtures.TextBody("横");
            body.BodyProperties!.Vertical = A.TextVerticalValues.Horizontal;
            var path = ShapeWorkbookFixtures.CreateWorkbook(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, textBody: body));
            try
            {
                var shape = Assert.Single(_reader.ReadFile(path).Sheets[0].DrawingObjects.OfType<ShapeModel>());

                Assert.Null(shape.Text!.Vertical);
                Assert.Equal(0.0, shape.Text.RotationDegrees);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void セルの文字の回転を読み取る()
        {
            using var stream = File.OpenRead(Path.Combine(TestPaths.EdgeCaseSamplesRoot, "edge-order-form", "template.xlsx"));
            var sheet = Assert.Single(_reader.Read(stream, new WorkbookReadOptions(SheetNameFilter: "注文書")).Sheets);

            Assert.Equal(255, sheet.Cells[Core.CellAddress.Parse("A24")].Style.TextRotation);
            Assert.Equal(0, sheet.Cells[Core.CellAddress.Parse("B24")].Style.TextRotation);
        }
    }
}
