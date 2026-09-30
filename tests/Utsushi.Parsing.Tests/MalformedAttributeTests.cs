using System;
using System.IO;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 型付き属性に不正な文字列(例: <c>flipH="maybe"</c>、<c>r="abc"</c>)を持つファイルの扱い(要件6.4)。
    /// OpenXml SDK の <c>.Value</c> が投げる <see cref="FormatException"/>/<see cref="OverflowException"/> を、
    /// <see cref="InvalidExcelFileException"/>(<see cref="InvalidExcelFileReason.Corrupted"/>)に読み替えることを確認する。
    /// XMLとしては整形式のため、XMLパートの事前検査(要件6.7)では検出されず、読み取り中に初めて失敗する経路である。
    /// </summary>
    public sealed class MalformedAttributeTests
    {
        private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private readonly OpenXmlWorkbookReader _reader = new();

        // --- ワークシートの属性 ------------------------------------------------------

        [Theory]
        [InlineData("<row r=\"abc\"><c r=\"A1\" t=\"inlineStr\"><is><t>値</t></is></c></row>", typeof(FormatException))]
        [InlineData("<row r=\"99999999999\"><c r=\"A1\" t=\"inlineStr\"><is><t>値</t></is></c></row>", typeof(OverflowException))]
        [InlineData("<row r=\"-1\"><c r=\"A1\" t=\"inlineStr\"><is><t>値</t></is></c></row>", typeof(FormatException))] // 符号なし整数の負値は SDK では FormatException
        [InlineData("<row r=\"1\"><c r=\"A1\" s=\"x\" t=\"inlineStr\"><is><t>値</t></is></c></row>", typeof(FormatException))]
        [InlineData("<row r=\"1\"><c r=\"A1\" s=\"99999999999\" t=\"inlineStr\"><is><t>値</t></is></c></row>", typeof(OverflowException))]
        [InlineData("<row r=\"1\" hidden=\"maybe\"><c r=\"A1\" t=\"inlineStr\"><is><t>値</t></is></c></row>", typeof(FormatException))]
        public void ワークシートのセル行の属性値が不正ならCorruptedになる(string rowXml, Type innerType)
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var entry = SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path);
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    entry,
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<worksheet xmlns=\"" + MainNamespace + "\"><sheetData>" + rowXml + "</sheetData></worksheet>");

                AssertCorrupted(path, innerType);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("scale=\"abc\"", typeof(FormatException))]
        [InlineData("scale=\"-1\"", typeof(FormatException))]
        public void pageSetupの拡大縮小率が数値でなければCorruptedになる(string attribute, Type innerType)
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var entry = SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path);
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    entry,
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<worksheet xmlns=\"" + MainNamespace + "\">"
                    + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>値</t></is></c></row></sheetData>"
                    + "<pageSetup " + attribute + "/>"
                    + "</worksheet>");

                AssertCorrupted(path, innerType);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 図形(drawing)の属性 ----------------------------------------------------

        public static TheoryData<string, Type> MalformedDrawingAnchors() => new()
        {
            // 反転(boolean)
            { DrawingStyleWorkbookFixtures.ShapeAnchor(xfrmAttributes: "flipH=\"maybe\""), typeof(FormatException) },
            { DrawingStyleWorkbookFixtures.ShapeAnchor(xfrmAttributes: "flipV=\"maybe\""), typeof(FormatException) },

            // 色の変換(整数)
            {
                DrawingStyleWorkbookFixtures.ShapeAnchor(
                    spPrExtra: "<a:solidFill><a:srgbClr val=\"FF0000\"><a:lumMod val=\"abc\"/></a:srgbClr></a:solidFill>"),
                typeof(FormatException)
            },

            // スタイル参照の索引(整数)
            {
                DrawingStyleWorkbookFixtures.ShapeAnchor(
                    style: DrawingStyleWorkbookFixtures.ExcelDefaultShapeStyle.Replace(
                        "<a:lnRef idx=\"2\">", "<a:lnRef idx=\"x\">", StringComparison.Ordinal)),
                typeof(FormatException)
            },
        };

        [Theory]
        [MemberData(nameof(MalformedDrawingAnchors))]
        public void 図形の属性値が不正ならunsupportedElementsによらずCorruptedになる(string anchorXml, Type innerType)
        {
            var path = DrawingStyleWorkbookFixtures.CreateWorkbook(anchorXml, DrawingStyleWorkbookFixtures.OfficeThemeXml());
            try
            {
                AssertCorrupted(path, innerType);

                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error, ReportCode: "report")));
                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 属性値が正しい図形は同じ組み立て方で読み取れる()
        {
            // 上のテストが「組み立て方の誤り」ではなく「属性値の不正」で失敗していることを確かめる対照群。
            var path = DrawingStyleWorkbookFixtures.CreateWorkbook(
                DrawingStyleWorkbookFixtures.ShapeAnchor(
                    spPrExtra: "<a:solidFill><a:srgbClr val=\"FF0000\"><a:lumMod val=\"50000\"/></a:srgbClr></a:solidFill>",
                    style: DrawingStyleWorkbookFixtures.ExcelDefaultShapeStyle,
                    xfrmAttributes: "flipH=\"1\" flipV=\"0\""),
                DrawingStyleWorkbookFixtures.OfficeThemeXml());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.IsType<ShapeModel>(Assert.Single(sheet.DrawingObjects));
                Assert.True(shape.FlipHorizontal);
                Assert.False(shape.FlipVertical);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- ヘルパー ----------------------------------------------------------------

        private void AssertCorrupted(string path, Type innerType)
        {
            using var stream = File.OpenRead(path);
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "report")));

            Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.NotNull(ex.InnerException);
            Assert.IsType(innerType, ex.InnerException);
        }
    }
}
