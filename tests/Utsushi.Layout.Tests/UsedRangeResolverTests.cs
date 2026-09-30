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
    /// 印刷範囲が無いシートの使用範囲に描画オブジェクトを含める <see cref="UsedRangeResolver"/>(要件3.10、タスク27.3)の検証。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 寸法は計算しやすい値に揃える: 列幅 10(文字数単位)は最大数字幅7で 70px = 52.5pt、最大数字幅14で 140px = 105pt。
    /// 行高は 20pt。
    /// </para>
    /// <para>
    /// 前半は <see cref="UsedRangeResolver.GetOccupiedRange"/>・<see cref="UsedRangeResolver.Resolve"/> を直接、
    /// 後半は <see cref="ReportLayoutEngine.Compute"/> 経由で、印刷範囲の優先順位(帳票定義の上書き &gt; Excelの印刷範囲 &gt;
    /// 使用範囲)に組み込まれていることを確かめる。
    /// </para>
    /// </remarks>
    public sealed class UsedRangeResolverTests
    {
        private const double ColumnWidthPt = 52.5;
        private const double RowHeightPt = 20.0;

        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

        /// <summary>セルを持たない(値も書式も無い)シート。列幅10・行高20。</summary>
        private static SheetModel EmptySheet(PageSetupModel? pageSetup = null, params DrawingObjectModel[] drawingObjects) =>
            new(
                "テストシート",
                new Dictionary<CellAddress, CellModel>(),
                Array.Empty<MergedRange>(),
                Array.Empty<double>(),
                Array.Empty<double>(),
                10.0,
                RowHeightPt,
                new HashSet<int>(),
                new HashSet<int>(),
                pageSetup ?? NoMarginA4(),
                drawingObjects);

        /// <summary><paramref name="rows"/>×<paramref name="columns"/>(A1起点)にセルを持つシート。</summary>
        private static SheetModel SheetWithCells(int rows, int columns, params DrawingObjectModel[] drawingObjects)
        {
            var sheet = UniformSheet(rows, columns, columnWidth: 10.0, rowHeightPt: RowHeightPt, pageSetup: NoMarginA4());
            return sheet with { DrawingObjects = drawingObjects };
        }

        private static ShapeModel Shape(string anchorCell, PointPt anchorOffset, AnchorExtent extent) =>
            new(
                2u, ShapePresetType.Rect, Array.Empty<double>(), 0, null, new ShapeOutline(ArgbColor.Black, 1.0), null,
                CellAddress.Parse(anchorCell), anchorOffset, extent);

        private static ShapeModel FixedShape(string anchorCell, double widthPt, double heightPt, double offsetX = 0, double offsetY = 0) =>
            Shape(anchorCell, new PointPt(offsetX, offsetY), new FixedAnchorExtent(widthPt, heightPt));

        private static ShapeModel SpanShape(string anchorCell, string toCell, double toOffsetX, double toOffsetY) =>
            Shape(anchorCell, default, new CellSpanAnchorExtent(CellAddress.Parse(toCell), new PointPt(toOffsetX, toOffsetY)));

        private static ImageModel FixedImage(string anchorCell, double widthPt, double heightPt) =>
            new(3u, new byte[] { 0x89, 0x50, 0x4E, 0x47 }, "image/png", 0, CellAddress.Parse(anchorCell), default,
                new FixedAnchorExtent(widthPt, heightPt));

        // -- 二セルアンカー -----------------------------------------------------------------------

        [Theory]
        // 起点, 終端セル, 終端のオフセット(X, Y), 期待する範囲
        [InlineData("A1", "C3", 5.0, 5.0, "A1:C3")]    // 終端のセルに掛かっている
        [InlineData("A1", "C3", 0.0, 0.0, "A1:B2")]    // 終端がセルの境界ちょうど → 1つ手前のセルまで
        [InlineData("A1", "C3", 0.0, 5.0, "A1:B3")]    // 列方向だけ境界ちょうど
        [InlineData("A1", "C3", 5.0, 0.0, "A1:C2")]    // 行方向だけ境界ちょうど
        [InlineData("A1", "C3", 0.01, 0.01, "A1:C3")]  // わずかでも掛かれば含める
        [InlineData("B2", "D5", 0.0, 0.0, "B2:C4")]
        [InlineData("B2", "B2", 0.0, 0.0, "B2:B2")]    // 大きさ0(起点と終端が同じ境界)でも起点のセルは含む
        [InlineData("B2", "C3", 0.0, 0.0, "B2:B2")]    // ちょうど1セル分
        [InlineData("A1", "A1", 0.0, 0.0, "A1:A1")]    // 1つ手前が0列目・0行目にならない
        public void 二セルアンカーは終端のセルまでで終端のオフセットが0なら1つ手前までになる(
            string from, string to, double toOffsetX, double toOffsetY, string expected)
        {
            var sheet = EmptySheet();

            var range = UsedRangeResolver.GetOccupiedRange(sheet, SpanShape(from, to, toOffsetX, toOffsetY), 7.0);

            Assert.Equal(CellRange.Parse(expected), range);
        }

        // -- 固定サイズ(一セルアンカー) ------------------------------------------------------------

        [Theory]
        // 起点, 起点のオフセット(X, Y), 幅, 高さ, 期待する範囲(列幅 52.5pt・行高 20pt)
        [InlineData("A1", 0.0, 0.0, 52.5, 20.0, "A1:A1")]     // 幅・高さがちょうど1セル分 → 次のセルには掛からない
        [InlineData("A1", 0.0, 0.0, 52.6, 20.1, "A1:B2")]     // わずかに超える
        [InlineData("A1", 0.0, 0.0, 105.0, 40.0, "A1:B2")]    // ちょうど2セル分
        [InlineData("A1", 0.0, 0.0, 105.1, 40.1, "A1:C3")]
        [InlineData("A1", 0.0, 0.0, 200.0, 95.0, "A1:D5")]    // 200 = 52.5×3 + 42.5、95 = 20×4 + 15
        [InlineData("A1", 50.0, 19.0, 2.5, 1.0, "A1:A1")]     // オフセット込みでちょうど境界
        [InlineData("A1", 50.0, 19.0, 5.0, 5.0, "A1:B2")]     // オフセット込みで境界を超える
        [InlineData("C3", 0.0, 0.0, 1.0, 1.0, "C3:C3")]
        [InlineData("C3", 0.0, 0.0, 0.0, 0.0, "C3:C3")]       // 大きさ0
        public void 固定サイズは列幅と行高をたどって終端のセルが決まる(
            string anchor, double offsetX, double offsetY, double widthPt, double heightPt, string expected)
        {
            var sheet = EmptySheet();

            var range = UsedRangeResolver.GetOccupiedRange(
                sheet, FixedShape(anchor, widthPt, heightPt, offsetX, offsetY), 7.0);

            Assert.Equal(CellRange.Parse(expected), range);
        }

        [Fact]
        public void 固定サイズは列ごとに異なる列幅と行高をたどる()
        {
            // A=4(trunc((1024+18)/256×7)=28px=21pt), B=20(trunc((5120+18)/256×7)=140px=105pt), 以降は既定の10(52.5pt)。
            // 行: 1=10pt, 2=50pt, 以降は既定の20pt。
            var sheet = EmptySheet() with
            {
                ColumnWidths = new[] { 4.0, 20.0 },
                RowHeights = new[] { 10.0, 50.0 },
            };

            var colA = ExcelUnitConverter.ColumnWidthToPoints(4.0, 7.0);
            var colB = ExcelUnitConverter.ColumnWidthToPoints(20.0, 7.0);

            // A・Bの合計ちょうど(境界) → B まで。行は1・2の合計(60pt)ちょうど → 2行目まで。
            Assert.Equal(
                CellRange.Parse("A1:B2"),
                UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("A1", colA + colB, 60.0), 7.0));

            // さらに1pt → C・3行目に掛かる。
            Assert.Equal(
                CellRange.Parse("A1:C3"),
                UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("A1", colA + colB + 1.0, 61.0), 7.0));

            // Aだけの幅を少し超える → B(幅の広い列)の途中で止まる。
            Assert.Equal(
                CellRange.Parse("A1:B2"),
                UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("A1", colA + 1.0, 11.0), 7.0));
        }

        // -- 非表示の列・行(幅・高さ0として扱い、描画時の配置と一致させる) ---------------------------------

        [Fact]
        public void 固定サイズの図形の右側に非表示の列が挟まると非表示の列を幅0としてたどる()
        {
            // B2 に幅200ptの図形、C〜F列が非表示。表示される列は B(52.5)・G・H・I(各52.5)で、
            // 200 = 52.5×3 + 42.5 なので I 列に掛かる(非表示の列を幅52.5とみなすと E 列で止まってしまう)。
            var sheet = EmptySheet() with { HiddenColumns = new HashSet<int> { 3, 4, 5, 6 } };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("B2", 200.0, 10.0), 7.0);

            Assert.Equal(CellRange.Parse("B2:I2"), range);
        }

        [Fact]
        public void 固定サイズの図形の下側に非表示の行が挟まると非表示の行を高さ0としてたどる()
        {
            // B2 に高さ50ptの図形、3〜5行目が非表示。表示される行は 2(20)・6(20)・7(20)で、50 = 20×2 + 10 → 7行目。
            var sheet = EmptySheet() with { HiddenRows = new HashSet<int> { 3, 4, 5 } };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("B2", 10.0, 50.0), 7.0);

            Assert.Equal(CellRange.Parse("B2:B7"), range);
        }

        [Fact]
        public void 起点のセル自体が非表示の列や行でも幅0としてたどる()
        {
            // 起点 B2 の列Bと行2が非表示。描画時は C 列・3行目の左上に置かれる。
            var sheet = EmptySheet() with
            {
                HiddenColumns = new HashSet<int> { 2 },
                HiddenRows = new HashSet<int> { 2 },
            };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("B2", 60.0, 30.0), 7.0);

            // C(52.5) → D に掛かる。3行目(20) → 4行目に掛かる。
            Assert.Equal(CellRange.Parse("B2:D4"), range);
        }

        [Fact]
        public void 非表示の列が続いても4096回で打ち切られる()
        {
            var sheet = EmptySheet() with { HiddenColumns = new HashSet<int>(Enumerable.Range(1, 5000)) };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("A1", 10.0, 10.0), 7.0);

            Assert.Equal(1 + 4096, range.LastColumn);
        }

        [Fact]
        public void 非表示の列と行を挟む図形の終端のセルは描画時の配置と一致する()
        {
            // 図形の終端(描画時の配置 = PageCommandBuilder)が、使用範囲の最後の列・行の内側に入ること。
            var sheet = EmptySheet(null, FixedShape("B2", 200.0, 50.0)) with
            {
                HiddenColumns = new HashSet<int> { 3, 4, 5, 6 },
                HiddenRows = new HashSet<int> { 3, 4, 5 },
            };

            var page = Assert.Single(_engine.Compute(ReportModel.Create(Definition(), sheet)).Pages);

            Assert.Equal((2, 9), page.ColumnRange);
            Assert.Equal((2, 7), page.RowRange);

            // 印刷範囲 B2:I7 の表示される列は B・G・H・I(4×52.5 = 210pt)、行は 2・6・7(3×20 = 60pt)。
            var shape = Assert.Single(Shapes(page));
            Assert.Equal(0.0, shape.Rect.Left, 3);
            Assert.Equal(0.0, shape.Rect.Top, 3);
            Assert.Equal(200.0, shape.Rect.Right, 3);
            Assert.Equal(50.0, shape.Rect.Bottom, 3);
            Assert.InRange(shape.Rect.Right, ColumnWidthPt * 3, ColumnWidthPt * 4);
            Assert.InRange(shape.Rect.Bottom, RowHeightPt * 2, RowHeightPt * 3);
        }

        [Fact]
        public void 二セルアンカーの終端は非表示の列や行があっても番地のまま()
        {
            // 二セルアンカーは終端のセルが明示されているため、列幅・行高をたどらない。
            var sheet = EmptySheet() with
            {
                HiddenColumns = new HashSet<int> { 3, 4 },
                HiddenRows = new HashSet<int> { 3 },
            };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, SpanShape("B2", "E5", 5.0, 5.0), 7.0);

            Assert.Equal(CellRange.Parse("B2:E5"), range);
        }

        [Theory]
        // 最大数字幅, 期待する範囲(幅100ptの図形。最大数字幅7なら列幅52.5pt、14なら105pt)
        [InlineData(7.0, "A1:B1")]
        [InlineData(14.0, "A1:A1")]
        public void 固定サイズの列のたどり方は最大数字幅に従う(double maxDigitWidthPx, string expected)
        {
            var sheet = EmptySheet();

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("A1", 100.0, 10.0), maxDigitWidthPx);

            Assert.Equal(CellRange.Parse(expected), range);
        }

        [Fact]
        public void 幅0の列が続いても4096回で打ち切られる()
        {
            // 列1〜5000が幅0。打ち切りが無ければ5001列目(幅52.5pt)で止まるが、4096回たどった時点で打ち切る。
            var sheet = EmptySheet() with { ColumnWidths = Enumerable.Repeat(0.0, 5000).ToList() };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("A1", 10.0, 10.0), 7.0);

            Assert.Equal(1 + 4096, range.LastColumn);
            Assert.Equal(1, range.LastRow);
        }

        [Fact]
        public void 高さ0の行が続いても4096回で打ち切られる()
        {
            var sheet = EmptySheet() with { RowHeights = Enumerable.Repeat(0.0, 5000).ToList() };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("A1", 10.0, 10.0), 7.0);

            Assert.Equal(1 + 4096, range.LastRow);
            Assert.Equal(1, range.LastColumn);
        }

        [Fact]
        public void 打ち切りは起点から数えるので途中から始まる図形も4096回で止まる()
        {
            var sheet = EmptySheet() with { ColumnWidths = Enumerable.Repeat(0.0, 10000).ToList() };

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape("CV1", 10.0, 10.0), 7.0); // CV = 100列目

            Assert.Equal(100 + 4096, range.LastColumn);
        }

        [Fact]
        public void 固定サイズの終端はシートの右端と下端を超えない()
        {
            var sheet = EmptySheet();
            var lastCell = new CellAddress(CellAddress.MaxRow, CellAddress.MaxColumn).ToString();

            var range = UsedRangeResolver.GetOccupiedRange(sheet, FixedShape(lastCell, 1000.0, 1000.0), 7.0);

            Assert.Equal(new CellRange(CellAddress.MaxRow, CellAddress.MaxColumn, CellAddress.MaxRow, CellAddress.MaxColumn), range);
        }

        // -- Resolve(セルの使用範囲との合成) -------------------------------------------------------

        [Fact]
        public void セルも描画オブジェクトも無ければnull()
        {
            Assert.Null(UsedRangeResolver.Resolve(EmptySheet(), 7.0));
        }

        [Fact]
        public void 描画オブジェクトが無ければセルの使用範囲と同じ()
        {
            var sheet = SheetWithCells(3, 2);

            Assert.Equal(sheet.GetUsedRange(), UsedRangeResolver.Resolve(sheet, 7.0));
        }

        [Fact]
        public void 表の右下に置いた図形の範囲まで使用範囲が広がる()
        {
            // セルは A1:B3。図形は E10 起点で 2列×2行弱 → E10:F11。
            var sheet = SheetWithCells(3, 2, FixedShape("E10", 60.0, 30.0));

            Assert.Equal(CellRange.Parse("A1:F11"), UsedRangeResolver.Resolve(sheet, 7.0));
        }

        [Fact]
        public void 表の左上に置いた図形の範囲まで使用範囲が広がる()
        {
            // セルは C3:D4 だけ。図形は A1:A1。
            var cells = new Dictionary<CellAddress, CellModel>();
            foreach (var address in new[] { "C3", "D4" })
            {
                var cell = CellAddress.Parse(address);
                cells[cell] = new CellModel(address, CellValueKind.Text, CellStyle.Default, address);
            }

            var sheet = EmptySheet(null, FixedShape("A1", 10.0, 10.0)) with { Cells = cells };

            Assert.Equal(CellRange.Parse("A1:D4"), UsedRangeResolver.Resolve(sheet, 7.0));
        }

        [Fact]
        public void 表の内側に収まる図形は使用範囲を変えない()
        {
            var sheet = SheetWithCells(5, 5, FixedShape("B2", 10.0, 10.0));

            Assert.Equal(CellRange.Parse("A1:E5"), UsedRangeResolver.Resolve(sheet, 7.0));
        }

        [Fact]
        public void 画像と図形の両方の範囲を合わせる()
        {
            // セル無し。画像は B2:B2、図形は D5 起点の2セルアンカーで G8 の境界まで → D5:F7。
            var sheet = EmptySheet(null, FixedImage("B2", 10.0, 10.0), SpanShape("D5", "G8", 0.0, 0.0));

            Assert.Equal(CellRange.Parse("B2:F7"), UsedRangeResolver.Resolve(sheet, 7.0));
        }

        [Fact]
        public void 結合セルの範囲も従来どおり使用範囲に含まれる()
        {
            var sheet = EmptySheet(null, FixedShape("A1", 10.0, 10.0)) with
            {
                MergedRanges = new[] { new MergedRange(CellRange.Parse("C3:E6")) },
            };

            Assert.Equal(CellRange.Parse("A1:E6"), UsedRangeResolver.Resolve(sheet, 7.0));
        }

        // -- ReportLayoutEngine 経由(印刷範囲の優先順位への組み込み) ------------------------------------

        [Fact]
        public void 印刷範囲が無ければ表の外に置いた図形まで出力される()
        {
            var sheet = SheetWithCells(3, 2, FixedShape("E10", 60.0, 30.0));

            var page = Assert.Single(_engine.Compute(ReportModel.Create(Definition(), sheet)).Pages);

            Assert.Equal((1, 11), page.RowRange);
            Assert.Equal((1, 6), page.ColumnRange);
            var shape = Assert.Single(Shapes(page));
            Assert.Equal(ColumnWidthPt * 4, shape.Rect.Left, 3);
            Assert.Equal(RowHeightPt * 9, shape.Rect.Top, 3);
        }

        [Fact]
        public void 図形だけのシートでも印刷範囲が決まり図形が出力される()
        {
            // 以前は使用範囲がセルだけから決まり「印刷対象のセルがありません」(LayoutComputationException)になっていた。
            var sheet = EmptySheet(null, FixedShape("B3", 60.0, 30.0));

            var page = Assert.Single(_engine.Compute(ReportModel.Create(Definition(), sheet)).Pages);

            Assert.Equal((3, 4), page.RowRange);
            Assert.Equal((2, 3), page.ColumnRange);
            var shape = Assert.Single(Shapes(page));
            Assert.Equal(0.0, shape.Rect.Left, 3);
            Assert.Equal(0.0, shape.Rect.Top, 3);
            Assert.Equal(60.0, shape.Rect.Width, 3);
            Assert.Equal(30.0, shape.Rect.Height, 3);
        }

        [Fact]
        public void 画像だけのシートでも印刷範囲が決まり画像が出力される()
        {
            var sheet = EmptySheet(null, FixedImage("C2", 10.0, 10.0));

            var page = Assert.Single(_engine.Compute(ReportModel.Create(Definition(), sheet)).Pages);

            Assert.Equal((2, 2), page.RowRange);
            Assert.Equal((3, 3), page.ColumnRange);
            Assert.Single(Images(page));
        }

        [Fact]
        public void セルも描画オブジェクトも無いシートは従来どおり印刷対象が無いエラーになる()
        {
            var ex = Assert.Throws<LayoutComputationException>(
                () => _engine.Compute(ReportModel.Create(Definition(), EmptySheet())));

            Assert.Contains("印刷対象のセルがありません", ex.Message, StringComparison.Ordinal);
        }

        [Theory]
        // 帳票定義の最大数字幅, 期待する列の範囲(幅100ptの図形のみ)
        [InlineData(7.0, 2)]
        [InlineData(14.0, 1)]
        public void 使用範囲の計算には帳票定義の最大数字幅が使われる(double maxDigitWidthPx, int expectedLastColumn)
        {
            var sheet = EmptySheet(null, FixedShape("A1", 100.0, 10.0));

            var page = Assert.Single(_engine.Compute(ReportModel.Create(Definition(maxDigitWidthPx), sheet)).Pages);

            Assert.Equal((1, expectedLastColumn), page.ColumnRange);
        }

        [Fact]
        public void Excelの印刷範囲があれば描画オブジェクトの範囲は使われない()
        {
            var pageSetup = NoMarginA4(printAreas: new[] { CellRange.Parse("A1:B3") });
            var sheet = SheetWithCells(3, 2, FixedShape("E10", 60.0, 30.0)) with { PageSetup = pageSetup };

            var page = Assert.Single(_engine.Compute(ReportModel.Create(Definition(), sheet)).Pages);

            Assert.Equal((1, 3), page.RowRange);
            Assert.Equal((1, 2), page.ColumnRange);
        }

        [Fact]
        public void 帳票定義のprintAreaがあれば描画オブジェクトの範囲は使われない()
        {
            var sheet = SheetWithCells(3, 2, FixedShape("E10", 60.0, 30.0));

            var page = Assert.Single(
                _engine.Compute(ReportModel.Create(Definition(printAreaOverride: CellRange.Parse("A1:B2")), sheet)).Pages);

            Assert.Equal((1, 2), page.RowRange);
            Assert.Equal((1, 2), page.ColumnRange);
        }

        [Fact]
        public void 帳票定義のprintAreaはExcelの印刷範囲より優先され描画オブジェクトの範囲も使われない()
        {
            var pageSetup = NoMarginA4(printAreas: new[] { CellRange.Parse("A1:B3") });
            var sheet = SheetWithCells(3, 2, FixedShape("E10", 60.0, 30.0)) with { PageSetup = pageSetup };

            var page = Assert.Single(
                _engine.Compute(ReportModel.Create(Definition(printAreaOverride: CellRange.Parse("A1:A1")), sheet)).Pages);

            Assert.Equal((1, 1), page.RowRange);
            Assert.Equal((1, 1), page.ColumnRange);
        }

        [Fact]
        public void 二セルアンカーの終端は始点から4096セルまでに抑える()
        {
            // 回帰テスト(code-reviewer指摘): 描画時の走査の上限と揃え、XFD1048576 まで伸びる図形1つで
            // 印刷範囲のセル数の上限(要件6.9)を超えて変換できなくなることを防ぐ。
            var sheet = LayoutFixtures.UniformSheet(rows: 1, columns: 1);
            var image = new ImageModel(
                1u, Array.Empty<byte>(), "image/png", 0,
                CellAddress.Parse("A1"), default,
                new CellSpanAnchorExtent(new CellAddress(CellAddress.MaxRow, CellAddress.MaxColumn), new PointPt(1, 1)));

            var range = UsedRangeResolver.GetOccupiedRange(sheet, image, 7.0);

            Assert.Equal(1 + 4096, range.LastRow);
            Assert.Equal(1 + 4096, range.LastColumn);
        }

        [Fact]
        public void 固定サイズの寸法は描画時と同じ5000ptまでに抑える()
        {
            // 行高 20pt の行を 5000pt 分(250行)たどった位置で止まり、1万pt の図形でも使用範囲は倍にならない。
            var sheet = LayoutFixtures.UniformSheet(rows: 1, columns: 1, rowHeightPt: 20.0);
            var image = new ImageModel(
                1u, Array.Empty<byte>(), "image/png", 0,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(1.0, 10000.0));

            var range = UsedRangeResolver.GetOccupiedRange(sheet, image, 7.0);

            Assert.Equal(250, range.LastRow);
        }
    }
}
