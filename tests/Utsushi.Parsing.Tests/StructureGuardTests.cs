using System;
using System.IO;
using System.Linq;
using System.Text;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 入力の構造に対する安全弁(要件6.7〜6.10、タスク25)の検証。
    /// <list type="bullet">
    /// <item>要件6.7: XMLパートの入れ子の深さ・大きさの上限、壊れたXML・DTDの拒否</item>
    /// <item>要件6.8: <c>&lt;col&gt;</c> の展開回数・改ページ件数・印刷範囲の個数の上限</item>
    /// <item>要件6.10: 画像の <c>r:embed</c> の参照先パートが無い場合の扱い</item>
    /// </list>
    /// 上限値を引数に取る <c>internal</c> ヘルパーで境界の両側を確認し、実際のリーダーを使う結合テストで
    /// 既定の上限値が読み取り経路に組み込まれていることを確認する。
    /// (要件6.9 の Layout 側は <c>Utsushi.Layout.Tests.LayoutSafetyLimitTests</c>)
    /// </summary>
    public sealed class StructureGuardTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        // --- 要件6.7: XMLの入れ子の深さ(ヘルパー) ------------------------------------

        [Theory]
        [InlineData(1, 1)]
        [InlineData(5, 5)]
        [InlineData(4, 5)]
        public void XMLの入れ子が上限段数以下なら例外にならない(int nestingLevels, int maxDepth)
        {
            using var xml = Utf8(NestedXml(nestingLevels));

            OpenXmlWorkbookReader.EnsureXmlWithinLimits(xml, maxDepth, maxBytes: 1024 * 1024, "/xl/test.xml", "report");
        }

        [Theory]
        [InlineData(2, 1)]
        [InlineData(6, 5)]
        [InlineData(100, 5)]
        public void XMLの入れ子が上限段数を1段でも超えるとTooLargeになる(int nestingLevels, int maxDepth)
        {
            using var xml = Utf8(NestedXml(nestingLevels));

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(xml, maxDepth, maxBytes: 1024 * 1024, "/xl/test.xml", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Contains("/xl/test.xml", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void XMLの入れ子は既定の上限段数ちょうどなら通り1段超えるとTooLargeになる()
        {
            const int max = OpenXmlWorkbookReader.MaxXmlElementDepth;

            using (var atLimit = Utf8(NestedXml(max)))
            {
                OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    atLimit, max, OpenXmlWorkbookReader.MaxXmlPartBytes, "/xl/test.xml", null);
            }

            using var overLimit = Utf8(NestedXml(max + 1));
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    overLimit, max, OpenXmlWorkbookReader.MaxXmlPartBytes, "/xl/test.xml", null));
            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
        }

        [Fact]
        public void 数千段にネストしたXMLでもスタックを使い切らずTooLargeになる()
        {
            // SDK の DOM 構築(再帰)ならスタックオーバーフローでプロセスごと落ちる規模(security-reviewer指摘の再現規模)。
            using var xml = Utf8(NestedXml(20_000));

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    xml, OpenXmlWorkbookReader.MaxXmlElementDepth, OpenXmlWorkbookReader.MaxXmlPartBytes, "/xl/test.xml", null));
            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
        }

        // --- 要件6.7: XMLパートの大きさ(ヘルパー) ------------------------------------

        [Fact]
        public void XMLパートの大きさが上限バイト数ちょうどなら例外にならない()
        {
            var bytes = Encoding.UTF8.GetBytes(NestedXml(3));
            using var xml = new MemoryStream(bytes);

            OpenXmlWorkbookReader.EnsureXmlWithinLimits(xml, maxDepth: 10, maxBytes: bytes.Length, "/xl/test.xml", "report");
        }

        [Fact]
        public void XMLパートの大きさが上限バイト数を1バイト超えるとTooLargeになる()
        {
            var bytes = Encoding.UTF8.GetBytes(NestedXml(3));
            using var xml = new MemoryStream(bytes);

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(
                    xml, maxDepth: 10, maxBytes: bytes.Length - 1, "/xl/test.xml", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal("report", ex.ReportCode);
        }

        // --- 要件6.7: 壊れたXML・DTD(ヘルパー) ----------------------------------------

        [Theory]
        [InlineData("<root><a></root>")] // 閉じタグの対応違い
        [InlineData("<root><a>")] // 閉じタグ欠け(途中で終わる)
        [InlineData("<root>")] // ルートの閉じタグ欠け
        [InlineData("not xml at all")]
        public void 壊れたXMLはCorruptedになる(string content)
        {
            using var xml = Utf8(content);

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(xml, maxDepth: 10, maxBytes: 1024, "/xl/test.xml", "report"));

            Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.IsType<System.Xml.XmlException>(ex.InnerException);
        }

        [Theory]
        [InlineData("<!DOCTYPE root><root/>")]
        [InlineData("<!DOCTYPE root [<!ENTITY e \"x\">]><root>&e;</root>")]
        [InlineData(
            "<!DOCTYPE root [<!ENTITY a \"aaaaaaaaaa\"><!ENTITY b \"&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;\">]><root>&b;</root>")] // Billion laughs の縮小版
        [InlineData("<!DOCTYPE root SYSTEM \"http://example.invalid/evil.dtd\"><root/>")] // 外部DTD
        public void DTDを含むXMLは拒否される(string content)
        {
            using var xml = Utf8(content);

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureXmlWithinLimits(xml, maxDepth: 10, maxBytes: 1024 * 1024, "/xl/test.xml", "report"));

            // DtdProcessing.Prohibit による XmlException が Corrupted に分類される。
            Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
        }

        // --- 要件6.7: 結合テスト(実際のリーダー) -------------------------------------

        [Fact]
        public void 図形グループを300段ネストしたdrawingはTooLargeになる()
        {
            var path = CreateWorkbookWithNestedGroups(nestedLevels: 300);
            try
            {
                // unsupportedElements の設定に関係なく、XMLの上限で読み取り前に落ちる。
                var ignoreEx = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ignoreEx.Reason);
                Assert.Contains("/xl/drawings/", ignoreEx.Message, StringComparison.Ordinal);

                using var stream = File.OpenRead(path);
                var errorEx = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error, ReportCode: "report")));
                Assert.Equal(InvalidExcelFileReason.TooLarge, errorEx.Reason);
                Assert.Equal("report", errorEx.ReportCode);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図形グループを100段ネストしたdrawingはXMLの上限では落ちずグループのネスト上限で扱われる()
        {
            // 100段はアプリ側の MaxShapeNestingDepth(5段)の範囲外だが、XMLの深さ上限(256段)の範囲内。
            Assert.True(100 > OpenXmlWorkbookReader.MaxShapeNestingDepth);

            var path = CreateWorkbookWithNestedGroups(nestedLevels: 100);
            try
            {
                using (var stream = File.OpenRead(path))
                {
                    var ex = Assert.Throws<UnsupportedWorkbookElementException>(
                        () => _reader.Read(stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error)));
                    Assert.Equal("GroupNestingTooDeep", ex.ElementKind);
                }

                // Ignore ならグループ全体が破棄され、読み取り自体は成功する。
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ワークシートのXMLの閉じタグが欠けているとCorruptedになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var entry = SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path);
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    entry,
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>値</t></is></c></row>");

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);

                // SDK の DOM 構築ではなく、XMLパートの事前検査で検出されている(パート名がメッセージに入る)。
                Assert.Contains(entry, ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ワークシートのXMLにDTDが含まれていると拒否される()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var entry = SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path);
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    entry,
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<!DOCTYPE worksheet [<!ENTITY e \"展開\">]>"
                    + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&e;</t></is></c></row></sheetData>"
                    + "</worksheet>");

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);

                // SDK の DOM 構築ではなく、XMLパートの事前検査で検出されている(パート名がメッセージに入る)。
                Assert.Contains(entry, ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 上限内の通常のワークブックはXMLの検査を通過して読み取れる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal("値", sheet.GetCell(CellAddress.Parse("A1"))?.DisplayValue);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 要件6.8: <col> の展開回数 ------------------------------------------------

        [Fact]
        public void 列定義の展開回数が上限以下なら例外にならない()
        {
            OpenXmlWorkbookReader.EnsureColumnExpansionsWithinLimit(expansions: 5, maxExpansions: 5, "シート1", "report");
        }

        [Fact]
        public void 列定義の展開回数が上限を超えるとTooLargeになる()
        {
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureColumnExpansionsWithinLimit(expansions: 6, maxExpansions: 5, "シート1", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal("report", ex.ReportCode);
            Assert.Contains("シート1", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 列定義の既定の上限は最大列数の4倍()
        {
            Assert.Equal(4 * CellAddress.MaxColumn, OpenXmlWorkbookReader.MaxColumnExpansionsPerSheet);
        }

        [Fact]
        public void 全列を覆う列定義を上限ちょうどまで重ねても読み取れる()
        {
            // 1〜16384列の <col> を4つ = 65,536回の展開(上限ちょうど)。
            var repeat = OpenXmlWorkbookReader.MaxColumnExpansionsPerSheet / CellAddress.MaxColumn;
            var path = SafetyLimitWorkbookFixtures.CreateWithColumns(
                Enumerable.Repeat((1U, (uint)CellAddress.MaxColumn), repeat));
            try
            {
                Assert.Single(_reader.ReadFile(path).Sheets);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 全列を覆う列定義の展開回数が上限を1超えるとTooLargeになる()
        {
            var repeat = OpenXmlWorkbookReader.MaxColumnExpansionsPerSheet / CellAddress.MaxColumn;
            var columns = Enumerable.Repeat((1U, (uint)CellAddress.MaxColumn), repeat).Append((1U, 1U));
            var path = SafetyLimitWorkbookFixtures.CreateWithColumns(columns);
            try
            {
                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 列の最大値を超えるmaxの列定義は最大列で打ち切って数える()
        {
            // max=20000 は最大列数(16384)に丸められてから数えるため、4つ並べても上限ちょうどに収まる。
            var repeat = OpenXmlWorkbookReader.MaxColumnExpansionsPerSheet / CellAddress.MaxColumn;
            var path = SafetyLimitWorkbookFixtures.CreateWithColumns(Enumerable.Repeat((1U, 20000U), repeat));
            try
            {
                Assert.Single(_reader.ReadFile(path).Sheets);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 要件6.8: 改ページの件数 --------------------------------------------------

        [Fact]
        public void 行の改ページが上限ちょうどなら読み取れる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithBreaks(OpenXmlWorkbookReader.MaxPageBreaksPerSheet, 0);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(OpenXmlWorkbookReader.MaxPageBreaksPerSheet, sheet.PageSetup.ManualRowBreaks.Count);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 行の改ページが上限を1超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithBreaks(OpenXmlWorkbookReader.MaxPageBreaksPerSheet + 1, 0);
            try
            {
                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "report")));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Equal("report", ex.ReportCode);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 列の改ページが上限ちょうどなら読み取れる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithBreaks(0, OpenXmlWorkbookReader.MaxPageBreaksPerSheet);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(OpenXmlWorkbookReader.MaxPageBreaksPerSheet, sheet.PageSetup.ManualColumnBreaks.Count);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 列の改ページが上限を1超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithBreaks(0, OpenXmlWorkbookReader.MaxPageBreaksPerSheet + 1);
            try
            {
                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 要件6.8: 印刷範囲の個数 --------------------------------------------------

        [Fact]
        public void 印刷範囲の個数が上限ちょうどなら読み取れる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithPrintAreas(OpenXmlWorkbookReader.MaxPrintAreasPerSheet);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(OpenXmlWorkbookReader.MaxPrintAreasPerSheet, sheet.PageSetup.PrintAreas.Count);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 印刷範囲の個数が上限を1超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithPrintAreas(OpenXmlWorkbookReader.MaxPrintAreasPerSheet + 1);
            try
            {
                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "report")));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Equal("report", ex.ReportCode);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 要件6.10: 参照先パートの無い画像 ----------------------------------------

        [Fact]
        public void 画像の参照先パートが無い場合unsupportedElementsがerrorならMissingImagePartになる()
        {
            var path = ShapeWorkbookFixtures.CreateWorkbook(SafetyLimitWorkbookFixtures.PictureAnchorWithMissingEmbed());
            try
            {
                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<UnsupportedWorkbookElementException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error, ReportCode: "report")));

                Assert.Equal("MissingImagePart", ex.ElementKind);
                Assert.Equal(ProcessingStage.Parsing, ex.Stage);
                Assert.Equal("report", ex.ReportCode);
                Assert.Equal("テストシート", ex.SheetName);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像の参照先パートが無い場合unsupportedElementsがignoreなら画像なしで読める()
        {
            // 同じ drawing 内の図形は読み取られ、壊れた画像だけが読み飛ばされる。
            var path = ShapeWorkbookFixtures.CreateWorkbook(
                SafetyLimitWorkbookFixtures.PictureAnchorWithMissingEmbed(),
                ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, id: 4U));
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);

                Assert.Empty(sheet.DrawingObjects.OfType<ImageModel>());
                Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- ヘルパー ----------------------------------------------------------------

        /// <summary><paramref name="levels"/> 段(ルートを含む)に入れ子になった要素のXML文字列。</summary>
        private static string NestedXml(int levels)
        {
            var builder = new StringBuilder(levels * 8);
            for (var i = 0; i < levels; i++)
            {
                builder.Append("<e>");
            }

            for (var i = 0; i < levels; i++)
            {
                builder.Append("</e>");
            }

            return builder.ToString();
        }

        private static MemoryStream Utf8(string content) => new(Encoding.UTF8.GetBytes(content));

        /// <summary>トップレベルのグループの下に <paramref name="nestedLevels"/> 段の <c>xdr:grpSp</c> を持つ .xlsx を作る。</summary>
        private static string CreateWorkbookWithNestedGroups(int nestedLevels)
        {
            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U);
            var nestedChain = ConnectorAndGroupWorkbookFixtures.WrapNestedGroups(nestedLevels, leaf);
            var topGroup = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new[] { nestedChain }, 0L, 0L, 900000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(topGroup);
            return ShapeWorkbookFixtures.CreateWorkbook(anchor);
        }
    }
}
