using System;
using System.IO;
using System.Text;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// パッケージ(ZIP)内のパートに対する安全弁(要件6.7)のうち、<see cref="StructureGuardTests"/> の後に追加したものの検証。
    /// <list type="bullet">
    /// <item>XMLパート1つの要素の数の上限 <c>MaxXmlElementsPerPart</c>(500万個)</item>
    /// <item>XMLパート1つの大きさの上限 <c>MaxXmlPartBytes</c>(64MiB)</item>
    /// <item>関係パート(<c>*.rels</c>)1つの関係の数の上限 <c>MaxRelationshipsPerPart</c>(1万件。<c>SpreadsheetDocument.Open</c> より前に検査)</item>
    /// <item>XMLパートの圧縮データが壊れている場合(<see cref="InvalidDataException"/>)を Corrupted に読み替えること</item>
    /// </list>
    /// </summary>
    public sealed class PackagePartGuardTests
    {
        private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string WorkbookRelsEntry = "xl/_rels/workbook.xml.rels";

        private readonly OpenXmlWorkbookReader _reader = new();

        // --- 要素の数(ヘルパー) -----------------------------------------------------

        [Theory]
        [InlineData(1, 1L)]
        [InlineData(5, 5L)]
        [InlineData(4, 5L)]
        public void XMLパートの要素の数が上限以下なら例外にならない(int elements, long maxElements)
        {
            using var xml = FlatXml(elements);

            OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                xml, maxDepth: 10, maxBytes: 1024 * 1024, maxElements, "/xl/test.xml", "report");
        }

        [Theory]
        [InlineData(2, 1L)]
        [InlineData(6, 5L)]
        [InlineData(100, 5L)]
        public void XMLパートの要素の数が上限を1超えるとTooLargeになる(int elements, long maxElements)
        {
            using var xml = FlatXml(elements);

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    xml, maxDepth: 10, maxBytes: 1024 * 1024, maxElements, "/xl/test.xml", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Contains("/xl/test.xml", ex.Message, StringComparison.Ordinal);
            Assert.Contains("要素の数", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 要素の数は深さに関係なく数える()
        {
            // 入れ子の要素(深さ3)も平らな要素と同じく1個として数える。
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes("<r><a><b/></a><c/></r>"));

            Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    xml, maxDepth: 10, maxBytes: 1024, maxElements: 3, "/xl/test.xml", null));
        }

        [Fact]
        public void 要素の数は既定の上限ちょうどなら通り1超えるとTooLargeになる()
        {
            // 要素数を引数に取らないオーバーロード(GuardXmlParts が使う)に既定の上限が組み込まれていることを確かめる。
            const long max = OpenXmlWorkbookReader.MaxXmlElementsPerPart;

            using (var atLimit = FlatXml(max))
            {
                OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    atLimit, OpenXmlWorkbookReader.MaxXmlElementDepth, OpenXmlWorkbookReader.MaxXmlPartBytes, "/xl/test.xml", null);
            }

            using var overLimit = FlatXml(max + 1);
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    overLimit, OpenXmlWorkbookReader.MaxXmlElementDepth, OpenXmlWorkbookReader.MaxXmlPartBytes, "/xl/test.xml", null));
            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Contains("要素の数", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 要素の数の上限はバイト数の上限より先に効く値になっている()
        {
            // 最小の要素(<e/> = 4バイト)を上限+1個並べても、パートの大きさの上限に達しない。
            // 逆だと、要素の数の上限は大きさの上限に隠れて意味を持たない。
            Assert.True((OpenXmlWorkbookReader.MaxXmlElementsPerPart + 1) * 4 < OpenXmlWorkbookReader.MaxXmlPartBytes);
            Assert.Equal(64L * 1024 * 1024, OpenXmlWorkbookReader.MaxXmlPartBytes);
        }

        // --- 要素の数・大きさ(結合テスト) ---------------------------------------------

        [Fact]
        public void ワークシートの要素の数が上限を1超えるとSDKがDOMを組み立てる前にTooLargeになる()
        {
            // worksheet + sheetData + row×(上限-2) + c + is + t = 上限+1個。
            var rows = (int)OpenXmlWorkbookReader.MaxXmlElementsPerPart - 2;
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var entry = SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path);
                var builder = new StringBuilder(rows * 6 + 256);
                builder.Append("<worksheet xmlns=\"").Append(MainNamespace).Append("\"><sheetData>");
                builder.Append("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>値</t></is></c></row>");
                builder.Insert(builder.Length, "<row/>", rows - 1);
                builder.Append("</sheetData></worksheet>");
                SafetyLimitWorkbookFixtures.ReplaceEntry(path, entry, builder.ToString());

                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "report")));

                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Equal("report", ex.ReportCode);
                Assert.Contains("要素の数", ex.Message, StringComparison.Ordinal);
                Assert.Contains(entry, ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ワークシートのXMLが64MiBを超えるとTooLargeになる()
        {
            // 要素は数個だけで、空白で大きさだけを上限+1バイトより大きくする。
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var entry = SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path);
                var padding = new string(' ', (int)OpenXmlWorkbookReader.MaxXmlPartBytes + 1);
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    entry,
                    "<worksheet xmlns=\"" + MainNamespace + "\"><sheetData>" + padding + "</sheetData></worksheet>");

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));

                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Contains("大きさ", ex.Message, StringComparison.Ordinal);
                Assert.Contains(entry, ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 関係の数(ヘルパー) -----------------------------------------------------

        [Theory]
        [InlineData(0, 5)]
        [InlineData(3, 5)]
        [InlineData(4, 5)]
        [InlineData(5, 5)] // 回帰テスト: 以前はルート要素 Relationships も数えており、上限ちょうどで TooLarge になった
        public void 関係の数が上限未満なら例外にならない(int relationships, int maxRelationships)
        {
            using var rels = RelsXml(relationships);

            OpenXmlWorkbookReader.EnsureRelationshipCountWithinLimit(rels, maxRelationships, "xl/_rels/test.xml.rels", "report");
        }

        [Theory]
        [InlineData(6, 5)]
        [InlineData(100, 5)]
        public void 関係の数が上限を1超えるとTooLargeになる(int relationships, int maxRelationships)
        {
            using var rels = RelsXml(relationships);

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureRelationshipCountWithinLimit(
                    rels, maxRelationships, "xl/_rels/test.xml.rels", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Contains("xl/_rels/test.xml.rels", ex.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("<Relationships><Relationship Id=\"rId1\"")] // 途中で終わる
        [InlineData("not xml at all")]
        [InlineData("<!DOCTYPE Relationships><Relationships/>")]
        public void 壊れた関係パートはこのヘルパーでは例外にせずOpen側の分類に委ねる(string content)
        {
            using var rels = new MemoryStream(Encoding.UTF8.GetBytes(content));

            OpenXmlWorkbookReader.EnsureRelationshipCountWithinLimit(rels, maxRelationships: 1, "xl/_rels/test.xml.rels", "report");
        }

        // --- 関係の数(結合テスト) -----------------------------------------------------

        [Fact]
        public void ワークブックの関係の数が上限を超えるとOpenより前にTooLargeになる()
        {
            // 既存のワークシートへの関係1件 + 外部ハイパーリンク 上限件 = 上限+1件。
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                SafetyLimitWorkbookFixtures.AddExternalRelationships(
                    path, WorkbookRelsEntry, OpenXmlWorkbookReader.MaxRelationshipsPerPart);

                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "report")));

                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Equal("report", ex.ReportCode);
                Assert.Contains(WorkbookRelsEntry, ex.Message, StringComparison.Ordinal);
                Assert.Null(ex.InnerException); // Open の失敗の読み替えではなく、事前検査で落ちている。
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ワークシートの関係パートも関係の数を検査する()
        {
            // workbook.xml.rels 以外の *.rels(ここではワークシートの drawing への関係)も対象になる。
            var path = ShapeWorkbookFixtures.CreateWorkbook(
                ShapeWorkbookFixtures.ShapeAnchor(DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle));
            try
            {
                var sheetRels = SafetyLimitWorkbookFixtures.EntryNameEndingWith(
                    path, "worksheets/_rels/" + SheetFileName(path) + ".rels");
                SafetyLimitWorkbookFixtures.AddExternalRelationships(path, sheetRels, OpenXmlWorkbookReader.MaxRelationshipsPerPart);

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));

                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Contains(sheetRels, ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ワークブックの関係の数が上限未満なら読み取れる()
        {
            // 既存の関係1件 + 外部ハイパーリンク(上限-2)件 = 上限-1件。事前検査が誤検出せず、Open 以降も通ることを確かめる。
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                SafetyLimitWorkbookFixtures.AddExternalRelationships(
                    path, WorkbookRelsEntry, OpenXmlWorkbookReader.MaxRelationshipsPerPart - 2);

                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal("値", sheet.GetCell(CellAddress.Parse("A1"))?.DisplayValue);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 壊れた関係パートは事前検査を素通りしOpen側でInvalidExcelFileExceptionになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    WorkbookRelsEntry,
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\"");

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));

                // 事前検査(TooLarge)ではなく、Open の失敗(破損)として読み替えられている。
                // 回帰テスト: 以前は Open の XmlException を分類できず Reason=Unknown になっていた。
                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
                Assert.NotNull(ex.InnerException);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 壊れた圧縮データ(結合テスト) ---------------------------------------------

        [Fact]
        public void ワークシートの圧縮データが壊れているとCorruptedになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var entry = SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path);
                SafetyLimitWorkbookFixtures.CorruptEntryCompressedData(path, entry);

                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "report")));

                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
                Assert.Equal(ProcessingStage.Parsing, ex.Stage);
                Assert.Equal("report", ex.ReportCode);
                Assert.Contains(entry, ex.Message, StringComparison.Ordinal);
                Assert.True(
                    ex.InnerException is InvalidDataException or IOException,
                    $"InnerException の型: {ex.InnerException?.GetType().FullName ?? "null"}");
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 関係パートの圧縮データが壊れていても事前検査は例外を漏らさずInvalidExcelFileExceptionになる()
        {
            // GuardPackageSize の関係の数の検査は、壊れた圧縮データ(InvalidDataException/IOException)を
            // 握りつぶして Open 側の分類に委ねる。SDK や .NET の例外がそのまま漏れないことを確かめる。
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                SafetyLimitWorkbookFixtures.CorruptEntryCompressedData(path, WorkbookRelsEntry);

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.NotEqual(InvalidExcelFileReason.TooLarge, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- ヘルパー ----------------------------------------------------------------

        /// <summary>ルートを含めてちょうど <paramref name="elements"/> 個の要素を持つ平らなXML(<c>&lt;r&gt;&lt;e/&gt;…&lt;/r&gt;</c>)。</summary>
        private static MemoryStream FlatXml(long elements)
        {
            var children = elements - 1;
            var bytes = new byte[3 + (children * 4) + 4];
            var position = 0;
            position = Write(bytes, position, "<r>");
            var child = Encoding.ASCII.GetBytes("<e/>");
            for (var i = 0L; i < children; i++)
            {
                Buffer.BlockCopy(child, 0, bytes, position, child.Length);
                position += child.Length;
            }

            Write(bytes, position, "</r>");
            return new MemoryStream(bytes);
        }

        private static int Write(byte[] buffer, int position, string ascii)
        {
            var bytes = Encoding.ASCII.GetBytes(ascii);
            Buffer.BlockCopy(bytes, 0, buffer, position, bytes.Length);
            return position + bytes.Length;
        }

        /// <summary><paramref name="relationships"/> 件の関係を持つ関係パート。</summary>
        private static MemoryStream RelsXml(int relationships)
        {
            var builder = new StringBuilder();
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (var i = 0; i < relationships; i++)
            {
                builder.Append("<Relationship Id=\"rId").Append(i)
                    .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\"")
                    .Append(" Target=\"https://example.invalid/\" TargetMode=\"External\"/>");
            }

            builder.Append("</Relationships>");
            return new MemoryStream(Encoding.UTF8.GetBytes(builder.ToString()));
        }

        /// <summary>最初のワークシートパートのファイル名(例: <c>sheet1.xml</c>)。</summary>
        private static string SheetFileName(string path) =>
            Path.GetFileName(SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path));
    }
}
