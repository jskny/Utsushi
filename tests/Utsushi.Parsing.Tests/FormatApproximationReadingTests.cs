using System.IO;
using System.Linq;
using Utsushi.Core;
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

        [Fact]
        public void 条件付き書式の範囲は空白で区切って読み有効な範囲より長い区切りは読み飛ばす()
        {
            var worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(
                new DocumentFormat.OpenXml.Spreadsheet.SheetData(),
                new DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                {
                    SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue>
                    {
                        InnerText = "  A1 " + new string('A', 1000) + "\t$XFD$1048575:$XFD$1048576  B2:C3 ",
                    },
                });

            var ranges = OpenXmlWorkbookReader.ReadConditionalFormattingRanges(worksheet, maxRanges: 100);

            Assert.Equal(new[] { "A1:A1", "XFD1048575:XFD1048576", "B2:C3" }, ranges.Select(r => r.ToString()).ToArray());
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

        [Fact]
        public void 同じ書式のランを大量に並べた共有文字列はリッチテキストとして読まない()
        {
            // 同じ書式のランは重複を除いて1件として保持する(セルの数×ランの数の比較にしない。security-reviewer指摘)。
            var runs = Enumerable.Range(0, 5000)
                .Select(_ => new DocumentFormat.OpenXml.Spreadsheet.Run(
                    new DocumentFormat.OpenXml.Spreadsheet.RunProperties(),
                    new DocumentFormat.OpenXml.Spreadsheet.Text("字")))
                .ToArray();
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(new DocumentFormat.OpenXml.Spreadsheet.SharedStringItem(runs));
            try
            {
                var cell = _reader.ReadFile(path).Sheets[0].Cells[CellAddress.Parse("A2")];

                Assert.False(cell.HasRichText);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("<x:i val=\"0\"/>", false)]
        [InlineData("<x:b val=\"maybe\"/>", true)]
        [InlineData("<x:sz val=\"abc\"/>", false)]
        [InlineData("<x:vertAlign val=\"superscript\"/>", true)]
        public void ランの書式の不正な値はファイルの破損として扱わない(string runProperties, bool expected)
        {
            var run = new DocumentFormat.OpenXml.Spreadsheet.Run(
                new DocumentFormat.OpenXml.Spreadsheet.RunProperties(
                    "<x:rPr xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + runProperties + "</x:rPr>"),
                new DocumentFormat.OpenXml.Spreadsheet.Text("文字"));
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(new DocumentFormat.OpenXml.Spreadsheet.SharedStringItem(run));
            try
            {
                var cell = _reader.ReadFile(path).Sheets[0].Cells[CellAddress.Parse("A2")];

                Assert.Equal(expected, cell.HasRichText);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("# ?/?", 0.5, NumberFormatFallback.UnsupportedFormat)]
        [InlineData("0.00E+00", 12345678.0, NumberFormatFallback.UnsupportedFormat)]
        [InlineData("[h]:mm", -0.5, NumberFormatFallback.UnrepresentableValue)]
        [InlineData("@", 1.0, NumberFormatFallback.None)]
        [InlineData("#,##0", 1234.0, NumberFormatFallback.None)]
        [InlineData(null, 1.0, NumberFormatFallback.None)]
        public void 表示形式で表示できず標準にした場合を理由とともに区別する(string? format, double value, NumberFormatFallback expected)
        {
            var result = NumberFormatting.FormatWithFallback(value, format);

            Assert.Equal(expected, result.Fallback);
            Assert.Equal(expected != NumberFormatFallback.None, result.FellBackToGeneral);
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
        [InlineData("<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FFFFFF\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"000000\"/></a:gs></a:gsLst><a:path path=\"rect\"/></a:gradFill>", "path=rect")]
        public void 直接指定の塗りつぶしを近似または省略した理由を読む(string fill, string expectedPrefix)
        {
            var shape = ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(spPrExtra: fill));

            Assert.StartsWith(expectedPrefix, shape.FillApproximation);
        }

        [Theory]
        [InlineData(null, "fontRef=minor", "文字")]
        [InlineData("<a:latin typeface=\"+mn-lt\"/>", "+mn-lt", "文字")]
        [InlineData("<a:latin typeface=\"Arial\"/>", "ea=-", "文字")]
        [InlineData("<a:latin typeface=\"Arial\"/><a:ea typeface=\"+mn-ea\"/>", "ea=+mn-ea", "文字")]
        [InlineData("<a:latin typeface=\"Arial\"/><a:ea typeface=\"MS Gothic\"/>", null, "文字")]
        [InlineData("<a:latin typeface=\"Arial\"/>", null, "Total")]
        public void スタイルやテーマの書体を指す図形の文字を読む(string? runFont, string? expected, string text = "文字")
        {
            var shape = ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(
                style: DrawingStyleWorkbookFixtures.ExcelDefaultShapeStyle,
                textBody: DrawingStyleWorkbookFixtures.TextBody(text, runFont)));

            Assert.Equal(expected, shape.Text!.StyleFont);
        }

        [Fact]
        public void 図形の文字のfontRefの不正な値はファイルの破損として扱わずそのまま返す()
        {
            // 診断にしか使わない値のため、型付きの値として読まない(code-reviewer指摘)。
            var shape = ReadSingleShape(DrawingStyleWorkbookFixtures.ShapeAnchor(
                style: DrawingStyleWorkbookFixtures.ExcelDefaultShapeStyle.Replace(
                    "idx=\"minor\"", "idx=\"bogus\"", System.StringComparison.Ordinal),
                textBody: DrawingStyleWorkbookFixtures.TextBody("文字")));

            Assert.Equal("fontRef=bogus", shape.Text!.StyleFont);
        }

        [Fact]
        public void 画像の反転を読み取る()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng(), rotationEmu: 60000);
            try
            {
                SafetyLimitWorkbookFixtures.ModifyEntry(
                    path,
                    SafetyLimitWorkbookFixtures.EntryNameEndingWith(path, "drawing1.xml"),
                    xml => xml.Replace("<a:xfrm ", "<a:xfrm flipH=\"1\" ", System.StringComparison.Ordinal));

                var image = Assert.Single(_reader.ReadFile(path).Sheets[0].DrawingObjects.OfType<ImageModel>());

                Assert.True(image.FlipHorizontal);
                Assert.False(image.FlipVertical);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像の反転の不正な値はファイルの破損として扱わず反転なしとする()
        {
            // 診断にしか使わない値のため、不正な値でファイル全体を変換できなくしない(security-reviewer指摘)。
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng(), rotationEmu: 60000);
            try
            {
                SafetyLimitWorkbookFixtures.ModifyEntry(
                    path,
                    SafetyLimitWorkbookFixtures.EntryNameEndingWith(path, "drawing1.xml"),
                    xml => xml.Replace("<a:xfrm ", "<a:xfrm flipH=\"maybe\" ", System.StringComparison.Ordinal));

                var image = Assert.Single(_reader.ReadFile(path).Sheets[0].DrawingObjects.OfType<ImageModel>());

                Assert.False(image.FlipHorizontal);
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
