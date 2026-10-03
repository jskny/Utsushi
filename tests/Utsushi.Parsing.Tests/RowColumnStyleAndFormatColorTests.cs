using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// セルが無い位置の書式(行・列・ブックの標準の書式。要件1.10)と、数値書式の色の指定(要件4.12)の読み取り。
    /// </summary>
    public sealed class RowColumnStyleAndFormatColorTests
    {
        private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static readonly ArgbColor Yellow = new(0xFF, 0xFF, 0xFF, 0x00);
        private static readonly ArgbColor Red = new(0xFF, 0xFF, 0x00, 0x00);

        private readonly OpenXmlWorkbookReader _reader = new();

        // --- 要件1.10: 行・列・ブックの標準の書式 ------------------------------------

        [Fact]
        public void セルが無い位置は行の書式を列の書式より優先する()
        {
            // 2行目は行全体に黄色の塗りつぶし(xf 1)、C列は列全体に赤の太字(xf 2)。
            var sheet = ReadSheet(
                "<cols><col min=\"3\" max=\"3\" width=\"10\" style=\"2\" customWidth=\"1\"/></cols>"
                + "<sheetData>"
                + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>A1</t></is></c></row>"
                + "<row r=\"2\" s=\"1\" customFormat=\"1\"/>"
                + "</sheetData>");

            Assert.Equal(Yellow, sheet.GetEffectiveStyle(CellAddress.Parse("B2")).BackgroundColor);
            Assert.Equal(Yellow, sheet.GetEffectiveStyle(CellAddress.Parse("C2")).BackgroundColor);

            var columnStyle = sheet.GetEffectiveStyle(CellAddress.Parse("C5"));
            Assert.True(columnStyle.Font.Bold);
            Assert.Equal(Red, columnStyle.Font.Color);

            // 行・列の書式の無い位置はブックの標準の書式(cellXfs の0番)。
            Assert.Equal("IPAGothic", sheet.GetEffectiveStyle(CellAddress.Parse("E5")).Font.Name);
        }

        [Fact]
        public void セルがあればセルの書式を使う()
        {
            var sheet = ReadSheet(
                "<sheetData><row r=\"2\" s=\"1\" customFormat=\"1\"><c r=\"A2\" s=\"0\" t=\"inlineStr\"><is><t>A2</t></is></c></row></sheetData>");

            Assert.True(sheet.GetEffectiveStyle(CellAddress.Parse("A2")).BackgroundColor.IsTransparent);
            Assert.Equal(Yellow, sheet.GetEffectiveStyle(CellAddress.Parse("B2")).BackgroundColor);
        }

        [Fact]
        public void customFormatの無い行の書式は使わない()
        {
            var sheet = ReadSheet("<sheetData><row r=\"2\" s=\"1\"/></sheetData>");

            Assert.Empty(sheet.RowStyles);
            Assert.True(sheet.GetEffectiveStyle(CellAddress.Parse("B2")).BackgroundColor.IsTransparent);
        }

        [Fact]
        public void 列の書式は範囲のまま持ち最大列まで展開しない()
        {
            var sheet = ReadSheet(
                "<cols><col min=\"1\" max=\"16384\" width=\"9\" style=\"1\"/></cols><sheetData/>");

            var range = Assert.Single(sheet.ColumnStyles);
            Assert.Equal(1, range.FirstColumn);
            Assert.Equal(CellAddress.MaxColumn, range.LastColumn);
            Assert.Equal(Yellow, sheet.GetEffectiveStyle(new CellAddress(100, CellAddress.MaxColumn)).BackgroundColor);

            // 行・列の書式は使用範囲を広げない(要件1.10補足)。
            Assert.Null(sheet.GetUsedRange());
        }

        [Fact]
        public void 重なる列の書式は列番号の小さい範囲を優先し昇順に並べる()
        {
            var sheet = ReadSheet(
                "<cols>"
                + "<col min=\"5\" max=\"6\" style=\"2\"/>"
                + "<col min=\"1\" max=\"5\" style=\"1\"/>"
                + "</cols><sheetData/>");

            Assert.Equal(new[] { (1, 5), (6, 6) }, sheet.ColumnStyles.Select(r => (r.FirstColumn, r.LastColumn)));
            Assert.Equal(Yellow, sheet.GetEffectiveStyle(CellAddress.Parse("E1")).BackgroundColor);
            Assert.True(sheet.GetEffectiveStyle(CellAddress.Parse("F1")).Font.Bold);
        }

        // --- 要件4.12: 数値書式の色 ----------------------------------------------------

        [Theory]
        [InlineData(-1000.0, "#,##0;[Red]-#,##0", 0xFFFF0000u)]
        [InlineData(-1000.0, "#,##0;[赤]-#,##0", 0xFFFF0000u)]
        [InlineData(1000.0, "[Blue]#,##0;[Red]-#,##0", 0xFF0000FFu)]
        [InlineData(1000.0, "[青]#,##0", 0xFF0000FFu)]
        [InlineData(-5.0, "[Color10]0;[Color3]0", 0xFFFF0000u)]
        [InlineData(5.0, "[Color10]0", 0xFF008000u)]
        [InlineData(-5.0, "[$-411][Red]0;[Red]-0", 0xFFFF0000u)]
        [InlineData(45383.0, "[Red]yyyy/m/d", 0xFFFF0000u)]
        public void 表示に使うセクションの色を返す(double value, string format, uint expected)
        {
            var color = NumberFormatter.ResolveColor(value, format);

            Assert.Equal(
                new ArgbColor((byte)(expected >> 24), (byte)(expected >> 16), (byte)(expected >> 8), (byte)expected),
                color);
        }

        [Theory]
        [InlineData(1000.0, "#,##0;[Red]-#,##0")]
        [InlineData(1000.0, "\"[Red]\"#,##0")]
        [InlineData(1000.0, "General")]
        [InlineData(1000.0, null)]
        [InlineData(1000.0, "[Color57]0")]
        [InlineData(1000.0, "[>100]0")]
        [InlineData(1000.0, "[Red]@")]
        public void 色の指定が無いセクションではnullを返す(double value, string? format)
        {
            Assert.Null(NumberFormatter.ResolveColor(value, format));
        }

        [Fact]
        public void 組み込み書式の負の数を赤で表示する()
        {
            Assert.Equal(Red, NumberFormatter.ResolveColor(-1000, NumberFormatter.GetBuiltInFormat(38)));
            Assert.Null(NumberFormatter.ResolveColor(1000, NumberFormatter.GetBuiltInFormat(38)));
        }

        [Fact]
        public void 数値セルに数値書式の色を持たせ置換で消す()
        {
            var sheet = ReadSheet(
                "<sheetData><row r=\"1\"><c r=\"A1\" s=\"3\"><v>-3000</v></c><c r=\"B1\" s=\"3\"><v>3000</v></c></row></sheetData>");

            var negative = sheet.Cells[CellAddress.Parse("A1")];
            Assert.Equal("-3,000", negative.DisplayValue);
            Assert.Equal(Red, negative.FormatColor);
            Assert.Null(sheet.Cells[CellAddress.Parse("B1")].FormatColor);

            Assert.Null(negative.WithText("差し込んだ値").FormatColor);
        }

        // --- ヘルパー -------------------------------------------------------------

        /// <summary>
        /// cellXfs: 0=標準(IPAGothic)、1=黄色の塗りつぶし、2=赤の太字、3=数値書式 <c>#,##0;[Red]-#,##0</c> のブックを作り、
        /// ワークシートの中身を <paramref name="worksheetInnerXml"/> にして読む。
        /// </summary>
        private SheetModel ReadSheet(string worksheetInnerXml)
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook((workbook, _) =>
            {
                var stylesPart = workbook.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(
                    new NumberingFormats(new NumberingFormat { NumberFormatId = 164U, FormatCode = "#,##0;[Red]-#,##0" }) { Count = 1U },
                    new Fonts(
                        new Font(new FontSize { Val = 11 }, new FontName { Val = "IPAGothic" }),
                        new Font(new Bold(), new FontSize { Val = 11 }, new Color { Rgb = "FFFF0000" }, new FontName { Val = "IPAGothic" }))
                    { Count = 2U },
                    new Fills(
                        new Fill(new PatternFill { PatternType = PatternValues.None }),
                        new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                        new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFFFFF00" } }))
                    { Count = 3U },
                    new Borders(new Border()) { Count = 1U },
                    new CellFormats(
                        new CellFormat { FontId = 0U, FillId = 0U },
                        new CellFormat { FontId = 0U, FillId = 2U, ApplyFill = true },
                        new CellFormat { FontId = 1U, FillId = 0U, ApplyFont = true },
                        new CellFormat { FontId = 0U, FillId = 0U, NumberFormatId = 164U, ApplyNumberFormat = true })
                    { Count = 4U });
                stylesPart.Stylesheet.Save();
            });

            try
            {
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path),
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + $"<worksheet xmlns=\"{MainNamespace}\">{worksheetInnerXml}</worksheet>");
                return Assert.Single(_reader.ReadFile(path).Sheets);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
