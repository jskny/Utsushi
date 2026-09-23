using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 入力ファイル全体の規模に対する安全弁(要件6.6)の検証。実際の上限値
    /// (<c>MaxXlsxPackageBytes</c>=1GiB、<c>MaxSharedStringCount</c>=20万件、
    /// <c>MaxCellsPerSheet</c>=50万個)は現実的なユニットテストでは大量のデータを
    /// 用意しないと到達できないため、比較・例外構築のロジックを分離した
    /// テスト可能なオーバーロード(<c>internal</c>、上限を明示的に指定できる)を使う。
    /// </summary>
    public sealed class InputSizeGuardTests
    {
        // --- 行番号・セル数(ReadSheet内の走査で使う比較ロジック) -------------------

        [Fact]
        public void 行番号が上限以下なら例外にならない()
        {
            OpenXmlWorkbookReader.EnsureRowIndexWithinLimit(rowIndex: 5, maxRows: 5, "シート1", "report");
        }

        [Fact]
        public void 行番号が上限を超えると例外になる()
        {
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureRowIndexWithinLimit(rowIndex: 6, maxRows: 5, "シート1", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
        }

        [Fact]
        public void セル数が上限以下なら例外にならない()
        {
            OpenXmlWorkbookReader.EnsureCellCountWithinLimit(cellCount: 5, maxCells: 5, "シート1", "report");
        }

        [Fact]
        public void セル数が上限を超えると例外になる()
        {
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureCellCountWithinLimit(cellCount: 6, maxCells: 5, "シート1", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
        }

        // --- シーク不可ストリームの複製時のバイト数上限 ------------------------------

        [Fact]
        public void シーク不可ストリームの複製は上限以下なら成功する()
        {
            using var source = new NonSeekableStream(new byte[] { 1, 2, 3, 4, 5 });
            using var destination = new MemoryStream();

            OpenXmlWorkbookReader.CopyWithSizeLimit(source, destination, maxBytes: 10, reportCode: "report");

            Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, destination.ToArray());
        }

        [Fact]
        public void シーク不可ストリームの複製は上限を超えると例外になる()
        {
            using var source = new NonSeekableStream(new byte[20]);
            using var destination = new MemoryStream();

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.CopyWithSizeLimit(source, destination, maxBytes: 10, reportCode: "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
        }

        // --- ZIP展開後サイズの事前ガード ---------------------------------------------

        [Fact]
        public void ZIPエントリの宣言サイズ合計が上限以下なら例外にならない()
        {
            using var zip = CreateZipWithEntrySize(50);

            OpenXmlWorkbookReader.GuardPackageSize(zip, maxBytes: 100, reportCode: "report");
            Assert.Equal(0, zip.Position);
        }

        [Fact]
        public void ZIPエントリの宣言サイズ合計が上限を超えると例外になる()
        {
            using var zip = CreateZipWithEntrySize(200);

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.GuardPackageSize(zip, maxBytes: 100, reportCode: "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
        }

        [Fact]
        public void 不正なZIP構造は例外にせずSpreadsheetDocumentOpen側の分類に委ねる()
        {
            using var garbage = new MemoryStream(new byte[] { 0x00, 0x01, 0x02, 0x03 });

            // InvalidDataExceptionを飲み込み、後続のSpreadsheetDocument.Openに判定を委ねる
            // (このメソッド自身は例外を送出しない)。
            OpenXmlWorkbookReader.GuardPackageSize(garbage, maxBytes: 1, reportCode: "report");
        }

        [Fact]
        public void シーク不可ストリームはZIPガードの対象外として何もしない()
        {
            using var source = new NonSeekableStream(new byte[100]);

            OpenXmlWorkbookReader.GuardPackageSize(source, maxBytes: 1, reportCode: "report");
        }

        private static MemoryStream CreateZipWithEntrySize(int uncompressedByteCount)
        {
            var stream = new MemoryStream();
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            {
                var entry = archive.CreateEntry("a.bin");
                using var entryStream = entry.Open();
                entryStream.Write(new byte[uncompressedByteCount], 0, uncompressedByteCount);
            }

            stream.Position = 0;
            return stream;
        }

        /// <summary>読み取り専用・非シーク(<see cref="Stream.CanSeek"/>=false)なストリームのテスト用ラッパー。</summary>
        private sealed class NonSeekableStream : MemoryStream
        {
            public NonSeekableStream(byte[] buffer)
                : base(buffer)
            {
            }

            public override bool CanSeek => false;
        }

        // --- 共有文字列の件数上限 ----------------------------------------------------

        [Fact]
        public void 共有文字列の数が上限以下なら例外にならない()
        {
            var path = CreateWorkbookWithSharedStrings(count: 3);
            try
            {
                using var document = SpreadsheetDocument.Open(path, isEditable: false);
                var result = OpenXmlWorkbookReader.ReadSharedStrings(document.WorkbookPart!, maxCount: 3, reportCode: "report");

                Assert.Equal(3, result.Count);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 共有文字列の数が上限を超えると例外になる()
        {
            var path = CreateWorkbookWithSharedStrings(count: 4);
            try
            {
                using var document = SpreadsheetDocument.Open(path, isEditable: false);

                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => OpenXmlWorkbookReader.ReadSharedStrings(document.WorkbookPart!, maxCount: 3, reportCode: "report"));

                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static string CreateWorkbookWithSharedStrings(int count)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-sharedstrings-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "テストシート" });

                var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                var table = new SharedStringTable();
                for (var i = 0; i < count; i++)
                {
                    table.Append(new SharedStringItem(new Text(i.ToString(CultureInfo.InvariantCulture))));
                }

                sharedStringPart.SharedStringTable = table;

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
                sharedStringPart.SharedStringTable.Save();
            }

            return path;
        }

        // --- スタイル要素(フォント・indexedColors)の件数上限(打ち切りのみ、例外化なし) ----

        [Fact]
        public void フォント数が上限を超えると超過分は読み取られず既定書式にフォールバックする()
        {
            const int maxEntries = 5;
            var workbookPart = CreateWorkbookPartWithFonts(fontCount: maxEntries + 3, cellFormatFontIds: new uint[] { 0, (uint)maxEntries });

            var colors = ColorResolver.Create(workbookPart);
            var styles = StyleTable.Create(workbookPart, colors, maxEntries);

            // cellXfs[0](fontId=0、上限内)は実際に設定した固有のフォント名を持つ。
            Assert.Equal("CustomFont0", styles.GetCellStyle(0).Font.Name);

            // cellXfs[1](fontId=maxEntries、上限で打ち切られた範囲外)は既定フォントにフォールバックする。
            Assert.Equal(FontStyle.Default.Name, styles.GetCellStyle(1).Font.Name);
        }

        [Fact]
        public void indexedColorsの数が上限を超えると超過分は読み取られずfallbackになる()
        {
            const int maxEntries = 5;
            var workbookPart = CreateWorkbookPartWithIndexedColors(colorCount: maxEntries + 3);

            var colors = ColorResolver.Create(workbookPart, maxEntries);
            var fallback = new ArgbColor(0xFF, 0x12, 0x34, 0x56);

            // 索引0(上限内)は実際に設定した固有の色を返す。
            var resolved = colors.Resolve(new X.Color { Indexed = 0U }, fallback);
            Assert.Equal(new ArgbColor(0xFF, 0x01, 0x02, 0x03), resolved);

            // 索引maxEntries(上限で打ち切られた範囲外)はfallbackになる。
            var truncated = colors.Resolve(new X.Color { Indexed = (uint)maxEntries }, fallback);
            Assert.Equal(fallback, truncated);
        }

        private static WorkbookPart CreateWorkbookPartWithFonts(int fontCount, IReadOnlyList<uint> cellFormatFontIds)
        {
            var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var fonts = new Fonts();
            for (var i = 0; i < fontCount; i++)
            {
                var name = i == 0 ? "CustomFont0" : "Font" + i.ToString(CultureInfo.InvariantCulture);
                fonts.Append(new Font(new FontName { Val = name }, new FontSize { Val = 11 }));
            }

            var cellFormats = new CellFormats();
            foreach (var fontId in cellFormatFontIds)
            {
                cellFormats.Append(new CellFormat { FontId = fontId });
            }

            stylesPart.Stylesheet = new Stylesheet(fonts, cellFormats);
            return workbookPart;
        }

        private static WorkbookPart CreateWorkbookPartWithIndexedColors(int colorCount)
        {
            var document = SpreadsheetDocument.Create(new MemoryStream(), SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var indexedColors = new IndexedColors();
            for (var i = 0; i < colorCount; i++)
            {
                var hex = i == 0 ? "FF010203" : "FF000000";
                indexedColors.Append(new RgbColor { Rgb = hex });
            }

            stylesPart.Stylesheet = new Stylesheet(new Colors(indexedColors));
            return workbookPart;
        }
    }
}
