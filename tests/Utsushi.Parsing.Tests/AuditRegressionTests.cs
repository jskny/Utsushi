using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using Dsf = Utsushi.Parsing.Tests.DrawingStyleWorkbookFixtures;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 監査で見つかった Parsing 層の不具合(H1/H3/H5/H6/M3〜M6・用紙コード・CellRange)の回帰テスト。
    /// </summary>
    public sealed class AuditRegressionTests
    {
        private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private readonly OpenXmlWorkbookReader _reader = new();

        // --- H1: c/@r・row/@r の省略 ------------------------------------------------

        [Fact]
        public void rが省略された行とセルは直前の行と直前のセルの次として読む()
        {
            var sheet = ReadSheetXml(
                "<sheetData>"
                + "<row r=\"2\"><c t=\"inlineStr\"><is><t>A2</t></is></c><c t=\"inlineStr\"><is><t>B2</t></is></c></row>"
                + "<row><c r=\"C3\" t=\"inlineStr\"><is><t>C3</t></is></c><c t=\"inlineStr\"><is><t>D3</t></is></c></row>"
                + "<row><c t=\"inlineStr\"><is><t>A4</t></is></c></row>"
                + "</sheetData>");

            Assert.Equal("A2", sheet.Cells[CellAddress.Parse("A2")].DisplayValue);
            Assert.Equal("B2", sheet.Cells[CellAddress.Parse("B2")].DisplayValue);
            Assert.Equal("C3", sheet.Cells[CellAddress.Parse("C3")].DisplayValue);
            Assert.Equal("D3", sheet.Cells[CellAddress.Parse("D3")].DisplayValue);
            Assert.Equal("A4", sheet.Cells[CellAddress.Parse("A4")].DisplayValue);
            Assert.Equal(5, sheet.Cells.Count);
        }

        [Fact]
        public void rを補った行番号にも行番号の上限を当てる()
        {
            var path = CreateWithSheetXml("<sheetData><row r=\"500000\"/><row/></sheetData>");
            try
            {
                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Contains("500001", ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- M4/M5: 行高・列幅・既定値 ------------------------------------------------

        [Fact]
        public void customHeightの無いhtも行高として採用する()
        {
            var sheet = ReadSheetXml("<sheetData><row r=\"1\" ht=\"30\"/><row r=\"2\" ht=\"20\" customHeight=\"1\"/></sheetData>");

            Assert.Equal(30.0, sheet.RowHeights[0]);
            Assert.Equal(20.0, sheet.RowHeights[1]);
        }

        [Theory]
        [InlineData("NaN")]
        [InlineData("INF")]
        [InlineData("-5")]
        [InlineData("1E+300")]
        public void 不正な行高は無視して既定の行高にする(string ht)
        {
            var sheet = ReadSheetXml($"<sheetData><row r=\"1\" ht=\"{ht}\" customHeight=\"1\"/></sheetData>");

            Assert.Equal(15.0, sheet.RowHeights[0]);
        }

        [Theory]
        [InlineData("NaN")]
        [InlineData("INF")]
        [InlineData("-5")]
        [InlineData("1E+300")]
        public void 不正な既定行高と既定列幅は既定値に戻す(string value)
        {
            var sheet = ReadSheetXml(
                $"<sheetFormatPr defaultRowHeight=\"{value}\" defaultColWidth=\"{value}\"/><sheetData><row r=\"1\"/></sheetData>");

            Assert.Equal(15.0, sheet.DefaultRowHeight);
            Assert.Equal(15.0, sheet.RowHeights[0]);
            Assert.True(double.IsNaN(sheet.DefaultColumnWidth)); // 暗黙の既定幅(Layout が baseColWidth から求める)
        }

        [Theory]
        [InlineData("NaN")]
        [InlineData("INF")]
        [InlineData("-5")]
        [InlineData("1E+300")]
        public void 不正な列幅は幅の指定が無いものとして扱う(string width)
        {
            var sheet = ReadSheetXml(
                $"<sheetFormatPr defaultRowHeight=\"15\" defaultColWidth=\"10\"/><cols><col min=\"1\" max=\"2\" width=\"{width}\" customWidth=\"1\"/></cols>"
                + "<sheetData/>");

            Assert.All(sheet.ColumnWidths, w => Assert.Equal(10.0, w));
        }

        [Fact]
        public void 正しい列幅と行高はそのまま読む()
        {
            var sheet = ReadSheetXml(
                "<cols><col min=\"1\" max=\"1\" width=\"255\" customWidth=\"1\"/></cols>"
                + "<sheetData><row r=\"1\" ht=\"409\"/><row r=\"2\" ht=\"0\"/></sheetData>");

            Assert.Equal(255.0, sheet.ColumnWidths[0]);
            Assert.Equal(409.0, sheet.RowHeights[0]);
            Assert.Equal(0.0, sheet.RowHeights[1]);
        }

        [Theory]
        [InlineData("NaN")]
        [InlineData("INF")]
        [InlineData("-5")]
        [InlineData("1E+300")]
        public void 不正な余白はその辺だけ既定値に戻す(string value)
        {
            var sheet = ReadSheetXml(
                "<sheetData/>"
                + $"<pageMargins left=\"{value}\" right=\"0.5\" top=\"{value}\" bottom=\"1\" header=\"{value}\" footer=\"0.2\"/>");

            var margins = sheet.PageSetup.Margins;
            Assert.Equal(Units.InchesToPoints(0.7), margins.LeftPt, 6);
            Assert.Equal(Units.InchesToPoints(0.5), margins.RightPt, 6);
            Assert.Equal(Units.InchesToPoints(0.75), margins.TopPt, 6);
            Assert.Equal(Units.InchesToPoints(1.0), margins.BottomPt, 6);
            Assert.Equal(Units.InchesToPoints(0.3), margins.HeaderPt, 6);
            Assert.Equal(Units.InchesToPoints(0.2), margins.FooterPt, 6);
        }

        [Theory]
        [InlineData("NaN")]
        [InlineData("INF")]
        [InlineData("-5")]
        [InlineData("0")]
        [InlineData("1E+300")]
        [InlineData("410")]
        public void 不正なフォントサイズは既定のサイズに戻す(string size)
        {
            var path = CreateWithFontSize(size);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(FontStyle.Default.SizePt, sheet.Cells[CellAddress.Parse("A1")].Style.Font.SizePt);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 範囲内のフォントサイズはそのまま読む()
        {
            var path = CreateWithFontSize("409");
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(409.0, sheet.Cells[CellAddress.Parse("A1")].Style.Font.SizePt);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- M3: インライン文字列のふりがな -----------------------------------------------

        [Fact]
        public void インライン文字列のリッチテキストにふりがなを混ぜない()
        {
            var sheet = ReadSheetXml(
                "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is>"
                + "<r><t>漢</t></r><r><rPr><b/></rPr><t>字</t></r>"
                + "<rPh sb=\"0\" eb=\"2\"><t>カンジ</t></rPh><phoneticPr fontId=\"0\"/>"
                + "</is></c></row></sheetData>");

            Assert.Equal("漢字", sheet.Cells[CellAddress.Parse("A1")].DisplayValue);
        }

        [Fact]
        public void インライン文字列のtとふりがなはtだけを読む()
        {
            var sheet = ReadSheetXml(
                "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is>"
                + "<t>東京</t><rPh sb=\"0\" eb=\"2\"><t>トウキョウ</t></rPh>"
                + "</is></c></row></sheetData>");

            Assert.Equal("東京", sheet.Cells[CellAddress.Parse("A1")].DisplayValue);
        }

        // --- H5: 手動改ページ ------------------------------------------------------

        [Fact]
        public void brkのidは0始まりでその行の上の改ページとして読む()
        {
            var sheet = ReadSheetXml(
                "<sheetData/>"
                + "<rowBreaks count=\"3\" manualBreakCount=\"3\">"
                + "<brk id=\"20\" max=\"16383\" man=\"1\"/><brk id=\"0\" max=\"16383\" man=\"1\"/><brk id=\"1048576\" max=\"16383\" man=\"1\"/>"
                + "</rowBreaks>"
                + "<colBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"3\" max=\"1048575\" man=\"1\"/></colBreaks>");

            // id=20 は 20 行目と 21 行目の間 = モデルでは「21 行目の手前」。id=0 と最終行より後ろは無視する。
            Assert.Equal(new[] { 21 }, sheet.PageSetup.ManualRowBreaks);
            Assert.Equal(new[] { 4 }, sheet.PageSetup.ManualColumnBreaks);
        }

        // --- H3/低7: 定義名 -------------------------------------------------------

        [Theory]
        [InlineData("Sheet1!$0:$1")]
        [InlineData("Sheet1!$1:$0")]
        [InlineData("Sheet1!$2000000:$2000000")]
        [InlineData("Sheet1!$1:$1048577")]
        public void 範囲外の行の印刷タイトルは無視する(string definition)
        {
            var titles = DefinedNameParser.ParsePrintTitles(definition);

            Assert.False(titles.HasRows);
        }

        [Fact]
        public void 範囲外の行の印刷タイトルを無視しても列の印刷タイトルは残す()
        {
            var titles = DefinedNameParser.ParsePrintTitles("Sheet1!$0:$1,Sheet1!$A:$B");

            Assert.False(titles.HasRows);
            Assert.Equal(1, titles.FirstColumn);
            Assert.Equal(2, titles.LastColumn);
        }

        [Fact]
        public void 最大行ちょうどの印刷タイトルは読む()
        {
            var titles = DefinedNameParser.ParsePrintTitles("Sheet1!$1048576:$1");

            Assert.Equal(1, titles.FirstRow);
            Assert.Equal(CellAddress.MaxRow, titles.LastRow);
        }

        [Fact]
        public void 範囲外の行の印刷タイトルを含むブックを読める()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook((workbook, _) =>
                workbook.Append(new DefinedNames(
                    new DefinedName("'" + SafetyLimitWorkbookFixtures.SheetName + "'!$0:$1")
                    {
                        Name = DefinedNameParser.PrintTitlesName,
                        LocalSheetId = 0U,
                    })));
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.False(sheet.PageSetup.PrintTitles.HasRows);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 印刷範囲は上限の1件超えで読むのを打ち切る()
        {
            var definition = string.Join(",", Enumerable.Range(1, 100).Select(i => "Sheet1!$A$" + i + ":$A$" + i));

            Assert.Equal(4, DefinedNameParser.ParsePrintArea(definition, 3).Count);
            Assert.Equal(100, DefinedNameParser.ParsePrintArea(definition).Count);
        }

        // --- H6: セル色のアルファ ---------------------------------------------------

        [Fact]
        public void セル色のrgbのアルファは無視して不透明にする()
        {
            var resolver = new ColorResolver(new uint[] { 0x00112233 }, Array.Empty<ArgbColor>());

            Assert.Equal(new ArgbColor(0xFF, 0xFF, 0x00, 0x00), resolver.Resolve(new Color { Rgb = "00FF0000" }, ArgbColor.Black));
            Assert.Equal(new ArgbColor(0xFF, 0x11, 0x22, 0x33), resolver.Resolve(new Color { Indexed = 0U }, ArgbColor.Black));
        }

        [Fact]
        public void アルファ00の色の文字と塗りを不透明として読む()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook((workbook, worksheet) =>
            {
                var stylesPart = workbook.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(
                    new Fonts(new Font(new FontSize { Val = 11 }, new Color { Rgb = "000000FF" })) { Count = 1U },
                    new Fills(
                        new Fill(new PatternFill { PatternType = PatternValues.None }),
                        new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
                        new Fill(new PatternFill(new ForegroundColor { Rgb = "0000FF00" }) { PatternType = PatternValues.Solid }))
                    { Count = 3U },
                    new Borders(new Border()) { Count = 1U },
                    new CellFormats(new CellFormat { FontId = 0U, FillId = 2U, ApplyFill = true }) { Count = 1U });
                stylesPart.Stylesheet.Save();
            });
            try
            {
                var style = Assert.Single(_reader.ReadFile(path).Sheets).Cells[CellAddress.Parse("A1")].Style;
                Assert.Equal(new ArgbColor(0xFF, 0x00, 0x00, 0xFF), style.Font.Color);
                Assert.Equal(new ArgbColor(0xFF, 0x00, 0xFF, 0x00), style.BackgroundColor);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- M6: 図形テキストの改行・フィールド・空の段落 -----------------------------------

        [Fact]
        public void 図形テキストの改行とフィールドと空の段落を読む()
        {
            var textBody =
                "<xdr:txBody><a:bodyPr/><a:lstStyle/>"
                + "<a:p><a:r><a:t>1行目</a:t></a:r><a:br><a:rPr sz=\"1400\"/></a:br><a:r><a:t>2行目</a:t></a:r></a:p>"
                + "<a:p><a:endParaRPr lang=\"ja-JP\"/></a:p>"
                + "<a:p><a:pPr algn=\"r\"/><a:fld id=\"{00000000-0000-0000-0000-000000000000}\" type=\"datetime1\"><a:rPr sz=\"900\"/><a:t>2026/10/1</a:t></a:fld></a:p>"
                + "</xdr:txBody>";

            var shape = ReadSingleShape(Dsf.ShapeAnchor(textBody: textBody));

            var paragraphs = shape.Text!.Paragraphs;
            Assert.Equal(3, paragraphs.Count);
            Assert.Equal(new[] { "1行目", "\n", "2行目" }, paragraphs[0].Runs.Select(r => r.Text));
            Assert.Equal(14.0, paragraphs[0].Runs[1].Font.SizePt);
            Assert.Empty(paragraphs[1].Runs);
            Assert.Equal("2026/10/1", Assert.Single(paragraphs[2].Runs).Text);
            Assert.Equal(9.0, paragraphs[2].Runs[0].Font.SizePt);
            Assert.Equal(HorizontalAlignment.Right, paragraphs[2].HAlign);
        }

        [Fact]
        public void 文字の無い図形テキストはテキストなしのままにする()
        {
            var textBody = "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"ja-JP\"/></a:p></xdr:txBody>";

            Assert.Null(ReadSingleShape(Dsf.ShapeAnchor(textBody: textBody)).Text);
        }

        [Theory]
        [InlineData(99)]
        [InlineData(400001)]
        [InlineData(-100)]
        public void 範囲外の図形の文字サイズは既定のサイズに戻す(int sz)
        {
            var shape = ReadSingleShape(Dsf.ShapeAnchor(textBody:
                "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz=\"" + sz + "\"/><a:t>x</a:t></a:r></a:p></xdr:txBody>"));

            Assert.Equal(FontStyle.Default.SizePt, shape.Text!.Paragraphs[0].Runs[0].Font.SizePt);
        }

        [Fact]
        public void 図形テキストの段落の区切りも文字数の上限に数える()
        {
            // 文字の無い段落だけを大量に並べても上限(2000文字)を迂回できない(Ignore モードでは図形ごと読み飛ばす)。
            var textBody = "<xdr:txBody><a:bodyPr/><a:lstStyle/>"
                + string.Concat(Enumerable.Repeat("<a:p/>", 2000))
                + "<a:p><a:r><a:t>x</a:t></a:r></a:p></xdr:txBody>";
            var path = Dsf.CreateWorkbook(Dsf.ShapeAnchor(textBody: textBody));
            try
            {
                Assert.Empty(Assert.Single(_reader.ReadFile(path).Sheets).DrawingObjects.OfType<ShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 段落1つで上限ちょうどの文字数の図形テキストは読む()
        {
            var textBody = "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>"
                + new string('x', 2000) + "</a:t></a:r></a:p></xdr:txBody>";

            var shape = ReadSingleShape(Dsf.ShapeAnchor(textBody: textBody));

            Assert.Equal(2000, shape.Text!.Paragraphs[0].Runs[0].Text.Length);
        }

        [Fact]
        public void 同じ画像パートを参照する画像はバイト列を共有する()
        {
            var path = ImageWorkbookFixtures.CreateWithManyPictures(3);
            try
            {
                var images = Assert.Single(_reader.ReadFile(path).Sheets).DrawingObjects.OfType<ImageModel>().ToList();

                Assert.Equal(3, images.Count);
                Assert.Same(images[0].Data, images[1].Data);
                Assert.Same(images[0].Data, images[2].Data);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 図形テキストの内側の余白(a:bodyPr の lIns 等) -------------------------------

        [Fact]
        public void 図形テキストの余白の指定が無ければnullにする()
        {
            var shape = ReadSingleShape(Dsf.ShapeAnchor(textBody:
                "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>x</a:t></a:r></a:p></xdr:txBody>"));

            Assert.Null(shape.Text!.Insets);
        }

        [Fact]
        public void 図形テキストの余白を読み無い辺と不正な辺は既定値にする()
        {
            // lIns=0、tIns=12700(1pt)、rIns は負(不正)、bIns は指定なし。
            var shape = ReadSingleShape(Dsf.ShapeAnchor(textBody:
                "<xdr:txBody><a:bodyPr lIns=\"0\" tIns=\"12700\" rIns=\"-12700\"/><a:lstStyle/>"
                + "<a:p><a:r><a:t>x</a:t></a:r></a:p></xdr:txBody>"));

            var insets = shape.Text!.Insets!;
            Assert.Equal(0.0, insets.LeftPt);
            Assert.Equal(1.0, insets.TopPt, 6);
            Assert.Equal(ShapeTextInsets.Default.RightPt, insets.RightPt);
            Assert.Equal(ShapeTextInsets.Default.BottomPt, insets.BottomPt);
        }

        // --- printOptions(ページ中央)・先頭ページ番号 ----------------------------------

        [Fact]
        public void ページ中央の指定と先頭ページ番号を読む()
        {
            var sheet = ReadSheetXml(
                "<sheetData/>"
                + "<printOptions horizontalCentered=\"1\" verticalCentered=\"true\"/>"
                + "<pageSetup paperSize=\"9\" firstPageNumber=\"5\" useFirstPageNumber=\"1\"/>");

            Assert.True(sheet.PageSetup.HorizontalCentered);
            Assert.True(sheet.PageSetup.VerticalCentered);
            Assert.Equal(5, sheet.PageSetup.FirstPageNumber);
        }

        [Theory]
        [InlineData("<pageSetup firstPageNumber=\"5\"/>")]
        [InlineData("<pageSetup firstPageNumber=\"5\" useFirstPageNumber=\"0\"/>")]
        [InlineData("<pageSetup firstPageNumber=\"abc\" useFirstPageNumber=\"1\"/>")]
        [InlineData("<pageSetup firstPageNumber=\"4294967295\" useFirstPageNumber=\"1\"/>")]
        [InlineData("<pageSetup useFirstPageNumber=\"1\"/>")]
        public void 先頭ページ番号が有効でなければ指定なしにする(string pageSetupXml)
        {
            var sheet = ReadSheetXml("<sheetData/>" + pageSetupXml);

            Assert.Null(sheet.PageSetup.FirstPageNumber);
            Assert.False(sheet.PageSetup.HorizontalCentered);
            Assert.False(sheet.PageSetup.VerticalCentered);
        }

        // --- 用紙コード ----------------------------------------------------------

        [Theory]
        [InlineData(12, 257, 364)]
        [InlineData(13, 182, 257)]
        [InlineData(43, 100, 148)]
        [InlineData(62, 182, 257)]
        [InlineData(69, 200, 148)]
        [InlineData(71, 240, 332)]
        [InlineData(73, 120, 235)]
        [InlineData(74, 90, 205)]
        [InlineData(79, 364, 257)]
        [InlineData(88, 128, 182)]
        public void 用紙コードはwingdiの定義に合わせる(int code, double widthMm, double heightMm)
        {
            var paper = PaperSizeTable.Resolve(code);

            Assert.Equal(code, paper.Code);
            Assert.Equal(PaperSize.FromMillimeters(code, paper.Name, widthMm, heightMm).WidthPt, paper.WidthPt, 6);
            Assert.Equal(PaperSize.FromMillimeters(code, paper.Name, widthMm, heightMm).HeightPt, paper.HeightPt, 6);
        }

        // --- CellRange -----------------------------------------------------------

        [Fact]
        public void CellRangeは最大行と最大列を超える範囲を作らない()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new CellRange(1, 1, CellAddress.MaxRow + 1, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => new CellRange(1, 1, 1, CellAddress.MaxColumn + 1));

            var whole = new CellRange(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn);
            Assert.Equal(new CellAddress(CellAddress.MaxRow, CellAddress.MaxColumn), whole.BottomRight);
        }

        // --- ヘルパー -------------------------------------------------------------

        private SheetModel ReadSheetXml(string worksheetInnerXml)
        {
            var path = CreateWithSheetXml(worksheetInnerXml);
            try
            {
                return Assert.Single(_reader.ReadFile(path).Sheets);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private ShapeModel ReadSingleShape(string anchorXml)
        {
            var path = Dsf.CreateWorkbook(anchorXml);
            try
            {
                return Assert.Single(Assert.Single(_reader.ReadFile(path).Sheets).DrawingObjects.OfType<ShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        /// <summary>ワークシートの中身(<c>worksheet</c> の子要素の XML)を差し替えた .xlsx を作る。</summary>
        private static string CreateWithSheetXml(string worksheetInnerXml)
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            SafetyLimitWorkbookFixtures.ReplaceEntry(
                path,
                SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path),
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + $"<worksheet xmlns=\"{MainNamespace}\">{worksheetInnerXml}</worksheet>");
            return path;
        }

        /// <summary>A1 のフォント(fontId=0)のサイズを <paramref name="size"/>(属性の文字列そのまま)にした .xlsx を作る。</summary>
        private static string CreateWithFontSize(string size)
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook((workbook, worksheet) =>
            {
                worksheet.Descendants<Cell>().First().StyleIndex = 0U;
                var stylesPart = workbook.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(
                    new Fonts(new Font(new FontSize { Val = 12345 }, new FontName { Val = "Arial" })) { Count = 1U },
                    new Fills(new Fill()) { Count = 1U },
                    new Borders(new Border()) { Count = 1U },
                    new CellFormats(new CellFormat { FontId = 0U }) { Count = 1U });
                stylesPart.Stylesheet.Save();
            });

            SafetyLimitWorkbookFixtures.ModifyEntry(
                path,
                SafetyLimitWorkbookFixtures.EntryNameEndingWith(path, "styles.xml"),
                xml => xml.Replace("\"12345\"", "\"" + size + "\"", StringComparison.Ordinal));
            return path;
        }
    }
}
