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
        public void 条件付き書式の範囲を拡張も含めて読み取る()
        {
            var worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                new DocumentFormat.OpenXml.Spreadsheet.SheetData(),
                new DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                {
                    SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1:B2 D4" },
                },
                new DocumentFormat.OpenXml.Spreadsheet.WorksheetExtensionList(
                    new DocumentFormat.OpenXml.Spreadsheet.WorksheetExtension(
                        new DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormattings(
                            new DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormatting(
                                new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence("F1:F9"))))
                    {
                        Uri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
                    }));

            var ranges = OpenXmlWorkbookReader.ReadConditionalFormattingRanges(worksheet, maxRanges: 100);

            Assert.Equal(new[] { "A1:B2", "D4:D4", "F1:F9" }, ranges.Select(r => r.ToString()).ToArray());
        }

        [Fact]
        public void 条件付き書式の範囲は上限で打ち切る()
        {
            var worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                new DocumentFormat.OpenXml.Spreadsheet.SheetData(),
                new DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                {
                    SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1 A2 A3 A4 A5" },
                },
                new DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                {
                    SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "B1" },
                });

            Assert.Equal(3, OpenXmlWorkbookReader.ReadConditionalFormattingRanges(worksheet, maxRanges: 3).Count);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, true)]
        public void ランの書式がセルの書式と異なるときだけリッチテキストとして読む(bool bold, bool expected)
        {
            var run = new DocumentFormat.OpenXml.Spreadsheet.Run(
                new DocumentFormat.OpenXml.Spreadsheet.RunProperties(
                    new DocumentFormat.OpenXml.Spreadsheet.RunFont { Val = FontStyle.Default.Name },
                    new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = FontStyle.Default.SizePt },
                    new DocumentFormat.OpenXml.Spreadsheet.Bold { Val = bold }),
                new DocumentFormat.OpenXml.Spreadsheet.Text("文字"));
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(new DocumentFormat.OpenXml.Spreadsheet.SharedStringItem(run));
            try
            {
                var cell = Assert.Single(_reader.ReadFile(path).Sheets[0].Cells.Values, c => c.DisplayValue == "文字");

                Assert.Equal(expected, cell.HasRichText);
                Assert.False(cell.WithText("書き込み").HasRichText);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("# ?/?", 0.5, true)]
        [InlineData("0.00E+00", 12345678.0, true)]
        [InlineData("[h]:mm", -0.5, true)]
        [InlineData("@", 1.0, false)]
        [InlineData("#,##0", 1234.0, false)]
        [InlineData(null, 1.0, false)]
        public void 表示形式で表示できず標準にした場合を区別する(string? format, double value, bool expected)
        {
            var result = NumberFormatting.FormatWithFallback(value, format);

            Assert.Equal(expected, result.FellBackToGeneral);
            Assert.Equal(NumberFormatting.Format(value, format), result.Text);
        }

        private ShapeModel ReadSingleShape(string anchorXml, string? themeXml = null)
        {
            var path = DrawingStyleWorkbookFixtures.CreateWorkbook(anchorXml, themeXml);
            try
            {
                return Assert.Single(_reader.ReadFile(path).Sheets[0].DrawingObjects.OfType<ShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static string Style(string fillIndex) =>
            DrawingStyleWorkbookFixtures.ExcelDefaultShapeStyle.Replace("<a:fillRef idx=\"1\">", "<a:fillRef idx=\"" + fillIndex + "\">");

        [Fact]
        public void 既定のテーマで単色のスタイルの塗りつぶしは近似としない()
        {
            Assert.Null(ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(style: Style("1"))).FillApproximation);
        }

        [Fact]
        public void 既定のテーマでグラデーションのスタイルの塗りつぶしは近似として読む()
        {
            Assert.Equal("fillRef idx=2", ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(style: Style("2"))).FillApproximation);
        }

        [Fact]
        public void テーマの塗りつぶしの書式が単色なら近似としない()
        {
            var shape = ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(style: Style("2")), DrawingStyleWorkbookFixtures.OfficeThemeXml());

            Assert.Null(shape.FillApproximation);
        }

        [Theory]
        [InlineData("<a:pattFill prst=\"pct5\"><a:fgClr><a:srgbClr val=\"000000\"/></a:fgClr><a:bgClr><a:srgbClr val=\"FFFFFF\"/></a:bgClr></a:pattFill>", "pattFill")]
        [InlineData("<a:grpFill/>", "grpFill")]
        [InlineData("<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FFFFFF\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"000000\"/></a:gs></a:gsLst><a:path path=\"rect\"/></a:gradFill>", "path=")]
        public void 直接指定の塗りつぶしを近似または省略した理由を読む(string fill, string expectedPrefix)
        {
            var shape = ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(spPrExtra: fill));

            Assert.StartsWith(expectedPrefix, shape.FillApproximation);
        }

        [Theory]
        [InlineData(null, "fontRef=Minor")]
        [InlineData("<a:latin typeface=\"+mn-lt\"/>", "+mn-lt")]
        [InlineData("<a:latin typeface=\"Arial\"/>", null)]
        public void スタイルやテーマの書体を指す図形の文字を読む(string? runFont, string? expected)
        {
            var shape = ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(
                style: DrawingStyleWorkbookFixtures.ExcelDefaultShapeStyle,
                textBody: DrawingStyleWorkbookFixtures.TextBody("文字", runFont)));

            Assert.Equal(expected, shape.Text!.StyleFont);
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
