using System;
using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using static Utsushi.Layout.Tests.LayoutFixtures;

namespace Utsushi.Layout.Tests
{
    /// <summary>
    /// レイアウト計算の規模に対する安全弁(要件6.9、タスク25)の検証。
    /// 印刷範囲の行数×列数の合計(<see cref="ReportLayoutEngine.MaxPrintRangeCells"/>)と
    /// 出力ページ数(<see cref="ReportLayoutEngine.MaxPagesPerDocument"/>)の上限を、
    /// 上限を引数に取るヘルパーで境界の両側を確認し、<see cref="ReportLayoutEngine.Compute"/> 経由の結合テストで
    /// 既定の上限が組み込まれていることを確認する。
    /// </summary>
    public sealed class LayoutSafetyLimitTests
    {
        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

        // --- 印刷範囲のセル数(ヘルパー) ---------------------------------------------

        public static IEnumerable<object[]> PrintRangeCellCasesWithinLimit() => new[]
        {
            // 範囲, 上限(行数×列数の合計がちょうど上限になる組み合わせを含む)
            new object[] { new[] { "A1:J10" }, 100L },
            new object[] { new[] { "A1:J10" }, 101L },
            new object[] { new[] { "A1:E10", "A20:E29" }, 100L },
            new object[] { Array.Empty<string>(), 0L },
        };

        public static IEnumerable<object[]> PrintRangeCellCasesOverLimit() => new[]
        {
            new object[] { new[] { "A1:J10" }, 99L },
            new object[] { new[] { "A1:E10", "A20:E29" }, 99L },
            new object[] { new[] { "A1:A1", "A1:A1" }, 1L },
        };

        [Theory]
        [MemberData(nameof(PrintRangeCellCasesWithinLimit))]
        public void 印刷範囲のセル数の合計が上限以下なら例外にならない(string[] ranges, long maxCells)
        {
            ReportLayoutEngine.EnsurePrintRangeCellsWithinLimit(
                ranges.Select(CellRange.Parse).ToList(), maxCells, "report", "シート1");
        }

        [Theory]
        [MemberData(nameof(PrintRangeCellCasesOverLimit))]
        public void 印刷範囲のセル数の合計が上限を超えるとLayoutComputationExceptionになる(string[] ranges, long maxCells)
        {
            var ex = Assert.Throws<LayoutComputationException>(
                () => ReportLayoutEngine.EnsurePrintRangeCellsWithinLimit(
                    ranges.Select(CellRange.Parse).ToList(), maxCells, "report", "シート1"));

            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Equal("シート1", ex.SheetName);
        }

        [Fact]
        public void 全セルの印刷範囲でも行数と列数の積がintを溢れずに判定される()
        {
            // 1,048,576 × 16,384 ≒ 1.7×10^10 は int を超える。long で積算していれば上限超過として検出される。
            var whole = new CellRange(1, 1, CellAddress.MaxRow, CellAddress.MaxColumn);

            Assert.Throws<LayoutComputationException>(
                () => ReportLayoutEngine.EnsurePrintRangeCellsWithinLimit(
                    new[] { whole }, ReportLayoutEngine.MaxPrintRangeCells, "report", "シート1"));
        }

        // --- 印刷範囲のセル数に印刷タイトルを含める(ヘルパー。security-reviewer指摘) ----
        // 印刷タイトルの行・列はすべての印刷範囲で繰り返し走査されるため、
        // 各印刷範囲について (行数 + タイトル行数) × (列数 + タイトル列数) を合計する。

        public static IEnumerable<object?[]> PrintRangeWithTitlesCasesWithinLimit() => new[]
        {
            // 範囲, タイトル先頭行, 末尾行, 先頭列, 末尾列, 上限(合計がちょうど上限)
            new object?[] { new[] { "A1:A1" }, 1, 9, null, null, 10L }, // (1+9)×1
            new object?[] { new[] { "A1:A1" }, null, null, 1, 9, 10L }, // 1×(1+9)
            new object?[] { new[] { "A1:B2" }, 1, 3, 1, 3, 25L }, // (2+3)×(2+3)
            new object?[] { new[] { "A1:A1", "C1:C1" }, 1, 4, null, null, 10L }, // 範囲ごとに足す: (1+4)×1 + (1+4)×1
            new object?[] { new[] { "A1:J10" }, null, null, null, null, 100L }, // タイトルなしは従来どおり
            new object?[] { new[] { "A1:A1" }, 1, 60000, 1, CellAddress.MaxColumn, 60001L * (CellAddress.MaxColumn + 1) },
        };

        public static IEnumerable<object?[]> PrintRangeWithTitlesCasesOverLimit() => new[]
        {
            new object?[] { new[] { "A1:A1" }, 1, 9, null, null, 9L },
            new object?[] { new[] { "A1:A1" }, null, null, 1, 9, 9L },
            new object?[] { new[] { "A1:B2" }, 1, 3, 1, 3, 24L },
            new object?[] { new[] { "A1:A1", "C1:C1" }, 1, 4, null, null, 9L },
        };

        [Theory]
        [MemberData(nameof(PrintRangeWithTitlesCasesWithinLimit))]
        public void 印刷タイトルを含めたセル数の合計が上限以下なら例外にならない(
            string[] ranges, int? firstRow, int? lastRow, int? firstColumn, int? lastColumn, long maxCells)
        {
            ReportLayoutEngine.EnsurePrintRangeCellsWithinLimit(
                ranges.Select(CellRange.Parse).ToList(),
                maxCells,
                "report",
                "シート1",
                new PrintTitles(firstRow, lastRow, firstColumn, lastColumn));
        }

        [Theory]
        [MemberData(nameof(PrintRangeWithTitlesCasesOverLimit))]
        public void 印刷タイトルを含めたセル数の合計が上限を1超えるとLayoutComputationExceptionになる(
            string[] ranges, int? firstRow, int? lastRow, int? firstColumn, int? lastColumn, long maxCells)
        {
            var ex = Assert.Throws<LayoutComputationException>(
                () => ReportLayoutEngine.EnsurePrintRangeCellsWithinLimit(
                    ranges.Select(CellRange.Parse).ToList(),
                    maxCells,
                    "report",
                    "シート1",
                    new PrintTitles(firstRow, lastRow, firstColumn, lastColumn)));

            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Equal("シート1", ex.SheetName);
        }

        [Fact]
        public void 小さな印刷範囲でも全列と6万行の印刷タイトルがあれば既定の上限を超える()
        {
            // 印刷範囲 A1:A1 だけなら1セルだが、タイトル 1:60000 行 × A:XFD 列を足すと約9.8億セル。
            var titles = new PrintTitles(1, 60000, 1, CellAddress.MaxColumn);

            Assert.Throws<LayoutComputationException>(
                () => ReportLayoutEngine.EnsurePrintRangeCellsWithinLimit(
                    new[] { CellRange.Parse("A1:A1") }, ReportLayoutEngine.MaxPrintRangeCells, "report", "シート1", titles));

            // タイトルを渡さなければ(従来の数え方)通ってしまう。
            ReportLayoutEngine.EnsurePrintRangeCellsWithinLimit(
                new[] { CellRange.Parse("A1:A1") }, ReportLayoutEngine.MaxPrintRangeCells, "report", "シート1");
        }

        // --- 出力ページ数(ヘルパー) -------------------------------------------------

        [Theory]
        [InlineData(0L, 5)]
        [InlineData(4L, 5)]
        [InlineData(5L, 5)]
        public void ページ数が上限以下なら例外にならない(long pageCount, int maxPages)
        {
            ReportLayoutEngine.EnsurePageCountWithinLimit(pageCount, maxPages, "report", "シート1");
        }

        [Theory]
        [InlineData(6L, 5)]
        [InlineData(5001L, 5000)]
        [InlineData(long.MaxValue, 5000)]
        public void ページ数が上限を超えるとLayoutComputationExceptionになる(long pageCount, int maxPages)
        {
            var ex = Assert.Throws<LayoutComputationException>(
                () => ReportLayoutEngine.EnsurePageCountWithinLimit(pageCount, maxPages, "report", "シート1"));

            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Equal("シート1", ex.SheetName);
        }

        // --- 結合テスト(Compute 経由) -----------------------------------------------

        [Fact]
        public void A1とXFD500000だけのシートは使用範囲が大きすぎてLayoutComputationExceptionになる()
        {
            // 2セルだけでも使用範囲は 500,000行 × 16,384列。SheetGrid を作る前に落ちること。
            var sheet = SparseSheet(
                new[] { CellAddress.Parse("A1"), CellAddress.Parse("XFD500000") },
                defaultRowHeightPt: 20.0,
                pageSetup: NoMarginA4());

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet));

            Assert.Equal("test-report", ex.ReportCode);
            Assert.Equal("テストシート", ex.SheetName);
            Assert.Contains("印刷範囲", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 帳票定義で指定した巨大な印刷範囲もLayoutComputationExceptionになる()
        {
            var sheet = SparseSheet(new[] { CellAddress.Parse("A1") }, defaultRowHeightPt: 20.0, pageSetup: NoMarginA4());

            var ex = Assert.Throws<LayoutComputationException>(
                () => Compute(sheet, Definition(printAreaOverride: CellRange.Parse("A1:XFD1048576"))));

            Assert.Contains("印刷範囲", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 複数の印刷範囲の合計が上限を超えるとLayoutComputationExceptionになる()
        {
            // 1範囲あたり 1,000行 × 1,000列 = 100万セル。2つで上限ちょうど、3つで超過。
            var range = CellRange.Parse("A1:ALL1000");
            Assert.Equal(1_000_000L, (long)range.RowCount * range.ColumnCount);

            var sheet = SparseSheet(
                new[] { CellAddress.Parse("A1") },
                defaultRowHeightPt: 20.0,
                pageSetup: NoMarginA4(printAreas: new[] { range, range, range }));

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet));
            Assert.Contains("印刷範囲", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 印刷範囲がA1だけでも巨大な印刷タイトルがあるとComputeでLayoutComputationExceptionになる()
        {
            // 印刷範囲 A1:A1 と、タイトル 1:60000 行 × A:XFD 列。SheetGrid を作る前に落ちること。
            var sheet = SparseSheet(
                new[] { CellAddress.Parse("A1") },
                defaultRowHeightPt: 20.0,
                pageSetup: NoMarginA4(
                    printAreas: new[] { CellRange.Parse("A1:A1") },
                    printTitles: new PrintTitles(1, 60000, 1, CellAddress.MaxColumn)));

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet));

            Assert.Equal("test-report", ex.ReportCode);
            Assert.Equal("テストシート", ex.SheetName);
            Assert.Contains("印刷範囲", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 上限内の印刷範囲でも印刷タイトルを足して上限を超えるとComputeでLayoutComputationExceptionになる()
        {
            // 印刷範囲 1,000行 × 1,000列 = 100万セル(単独では上限200万の内側)。
            // タイトル 1,001 行を足すと (1000 + 1001) × 1000 = 200.1万セルで上限を超える。
            var range = CellRange.Parse("A1:ALL1000");
            Assert.Equal(1000, range.ColumnCount);

            var sheet = SparseSheet(
                new[] { CellAddress.Parse("A1") },
                defaultRowHeightPt: 20.0,
                pageSetup: NoMarginA4(printAreas: new[] { range }, printTitles: new PrintTitles(1, 1001, null, null)));

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet));
            Assert.Contains("印刷範囲", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 通常の大きさの印刷タイトルはComputeで上限に掛からない()
        {
            // 見出し1行・1列を繰り返す一般的な帳票の設定で、誤って上限に掛からないことの確認。
            var sheet = UniformSheet(
                rows: 100,
                columns: 5,
                pageSetup: NoMarginA4(
                    printAreas: new[] { CellRange.Parse("A1:E100") },
                    printTitles: new PrintTitles(1, 1, 1, 1)));

            var layout = Compute(sheet);

            Assert.True(layout.Pages.Count > 1);
        }

        [Fact]
        public void 行数が多く出力ページ数が上限を1超えるとLayoutComputationExceptionになる()
        {
            // 行高800pt(A4縦・余白ゼロの印字可能高さ約842ptに1行しか入らない)× 上限+1行 = 上限+1ページ。
            var rows = ReportLayoutEngine.MaxPagesPerDocument + 1;
            var sheet = SparseSheet(
                new[] { CellAddress.Parse("A1") },
                defaultRowHeightPt: 800.0,
                pageSetup: NoMarginA4(printAreas: new[] { new CellRange(1, 1, rows, 1) }));

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet));

            Assert.Contains("ページ数", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 出力ページ数が上限ちょうどなら組み立てられる()
        {
            var rows = ReportLayoutEngine.MaxPagesPerDocument;
            var sheet = SparseSheet(
                new[] { CellAddress.Parse("A1") },
                defaultRowHeightPt: 800.0,
                pageSetup: NoMarginA4(printAreas: new[] { new CellRange(1, 1, rows, 1) }));

            var layout = Compute(sheet);

            Assert.Equal(ReportLayoutEngine.MaxPagesPerDocument, layout.Pages.Count);
        }

        [Fact]
        public void 複数の印刷範囲にまたがる通しのページ数が上限を超えるとLayoutComputationExceptionになる()
        {
            // 1範囲目で上限ちょうど(5000ページ)、2範囲目の1ページで通し5001ページになる。
            var rows = ReportLayoutEngine.MaxPagesPerDocument;
            var sheet = SparseSheet(
                new[] { CellAddress.Parse("A1") },
                defaultRowHeightPt: 800.0,
                pageSetup: NoMarginA4(printAreas: new[]
                {
                    new CellRange(1, 1, rows, 1),
                    CellRange.Parse("C1:C1"),
                }));

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet));
            Assert.Contains("ページ数", ex.Message, StringComparison.Ordinal);
        }

        // --- ヘルパー ----------------------------------------------------------------

        private PagedLayout Compute(SheetModel sheet, ReportDefinition? definition = null) =>
            _engine.Compute(ReportModel.Create(definition ?? Definition(), sheet));

        /// <summary>指定したセルだけを持ち、行高・列幅はすべて既定値のシート(大規模な範囲を少ないメモリで表す)。</summary>
        private static SheetModel SparseSheet(
            IEnumerable<CellAddress> addresses, double defaultRowHeightPt, PageSetupModel pageSetup)
        {
            var cells = addresses.ToDictionary(
                a => a,
                a => new CellModel(a.ToString(), CellValueKind.Text, CellStyle.Default, a.ToString()));

            return new SheetModel(
                "テストシート",
                cells,
                Array.Empty<MergedRange>(),
                Array.Empty<double>(),
                Array.Empty<double>(),
                10.0,
                defaultRowHeightPt,
                new HashSet<int>(),
                new HashSet<int>(),
                pageSetup,
                Array.Empty<DrawingObjectModel>());
        }
    }
}
