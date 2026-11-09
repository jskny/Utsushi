using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using static Utsushi.Layout.Tests.LayoutFixtures;

namespace Utsushi.Layout.Tests
{
    /// <summary>
    /// 書式の近似の記録(excel2pdf-consolidation 要件7、design.md 3.5)。印刷されるセル・要素だけを、診断のときだけ記録する。
    /// </summary>
    public sealed class FormatApproximationLayoutTests
    {
        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

        /// <summary>3行×5列(列幅10文字・行高20pt)の、指定セルだけを持つシート。印刷範囲は A1:C3。</summary>
        private static SheetModel Sheet(params (string Cell, CellModel Model)[] cells)
        {
            var map = cells.ToDictionary(c => CellAddress.Parse(c.Cell), c => c.Model);
            return UniformSheet(rows: 3, columns: 5, pageSetup: NoMarginA4(printAreas: new[] { CellRange.Parse("A1:C3") })) with
            {
                Cells = map,
            };
        }

        private static CellModel Text(string text, CellStyle style) => new(text, CellValueKind.Text, style, text);

        private PagedLayout Compute(SheetModel sheet, bool checkFit = true) =>
            _engine.Compute(ReportModel.Create(Definition(), sheet), checkFit);

        private static FormatApproximation Single(PagedLayout layout, FormatApproximationKind kind) =>
            Assert.Single(layout.FormatApproximations, a => a.Kind == kind);

        [Fact]
        public void 均等割り付けと選択範囲内で中央と縦の両端揃えを記録する()
        {
            var layout = Compute(Sheet(
                ("A1", Text("承認者", CellStyle.Default with { HAlign = HorizontalAlignment.Distributed })),
                ("B1", Text("受付日", CellStyle.Default with { HAlign = HorizontalAlignment.CenterContinuous })),
                ("C1", Text("備考", CellStyle.Default with { VAlign = VerticalAlignment.Justify }))));

            var horizontal = layout.FormatApproximations.Where(a => a.Kind == FormatApproximationKind.HorizontalAlignment).ToList();
            Assert.Equal(2, horizontal.Count);
            Assert.Contains(horizontal, a => a.Location == CellRange.Parse("A1:A1") && a.Detail == "horizontal=Distributed");
            Assert.Contains(horizontal, a => a.Location == CellRange.Parse("B1:B1") && a.Detail == "horizontal=CenterContinuous");
            Assert.Equal("vertical=Justify", Single(layout, FormatApproximationKind.VerticalAlignment).Detail);
            Assert.All(layout.FormatApproximations, a => Assert.Equal("テストシート", a.SheetName));
        }

        [Fact]
        public void 右揃えと中央揃えのセルのインデントを記録し左揃えは記録しない()
        {
            var layout = Compute(Sheet(
                ("A1", Text("右", CellStyle.Default with { HAlign = HorizontalAlignment.Right, Indent = 2 })),
                ("B1", Text("左", CellStyle.Default with { HAlign = HorizontalAlignment.Left, Indent = 2 }))));

            var indent = Single(layout, FormatApproximationKind.Indent);
            Assert.Equal(CellRange.Parse("A1:A1"), indent.Location);
            Assert.Equal("indent=2", indent.Detail);
        }

        [Fact]
        public void セルの文字の回転と縦書きを記録する()
        {
            var layout = Compute(Sheet(
                ("A1", Text("見出し", CellStyle.Default with { TextRotation = 90 })),
                ("B1", Text("縦書き", CellStyle.Default with { TextRotation = 255 }))));

            var rotations = layout.FormatApproximations.Where(a => a.Kind == FormatApproximationKind.TextRotation).ToList();
            Assert.Equal(2, rotations.Count);
            Assert.Contains(rotations, a => a.Location == CellRange.Parse("A1:A1") && a.Detail == "textRotation=90");
            Assert.Contains(rotations, a => a.Location == CellRange.Parse("B1:B1") && a.Detail == "textRotation=255" && a.Message.Contains("縦書き"));
        }

        [Fact]
        public void 図形の文字の縦書きを図形のアンカーと識別で記録する()
        {
            var text = new ShapeTextBody(
                new[] { new ShapeTextParagraph(new[] { new ShapeTextRun("縦", FontStyle.Default) }, HorizontalAlignment.Left) },
                VerticalAlignment.Top)
            {
                Vertical = "EastAsianVertical",
            };
            var shape = new ShapeModel(
                7u, ShapePresetType.Rect, System.Array.Empty<double>(), 0, null, null, text,
                CellAddress.Parse("B2"), new PointPt(0, 0), new FixedAnchorExtent(40.0, 30.0));
            var sheet = Sheet() with { DrawingObjects = new DrawingObjectModel[] { shape } };

            var rotation = Single(Compute(sheet), FormatApproximationKind.TextRotation);

            Assert.Equal(CellRange.Parse("B2:B2"), rotation.Location);
            Assert.Equal("vert=EastAsianVertical", rotation.Detail);
        }

        [Fact]
        public void 行の書式の網かけはセルが無い位置も印刷される範囲だけ記録する()
        {
            var patterned = CellStyle.Default with
            {
                BackgroundColor = new ArgbColor(0xFF, 0xEE, 0xEE, 0xEE),
                FillApproximation = FormatApproximationKind.CellFillPattern,
                FillApproximationDetail = "patternType=DarkTrellis",
            };
            var sheet = Sheet() with { RowStyles = new Dictionary<int, CellStyle> { [1] = patterned } };

            var fills = Compute(sheet).FormatApproximations.Where(a => a.Kind == FormatApproximationKind.CellFillPattern).ToList();

            // 印刷範囲 A1:C3 の1行目だけ(D1・E1 は印刷範囲の外)。
            Assert.Equal(new[] { "A1:A1", "B1:B1", "C1:C1" }, fills.Select(f => f.Location!.Value.ToString()).OrderBy(x => x).ToArray());
            Assert.All(fills, f => Assert.Equal("patternType=DarkTrellis", f.Detail));
        }

        private static ImageModel Image(uint id, double rotation, bool flip = false) =>
            new(id, new byte[] { 1 }, "image/png", rotation, CellAddress.Parse("B2"), new PointPt(0, 0), new FixedAnchorExtent(20.0, 10.0))
            {
                FlipHorizontal = flip,
            };

        [Theory]
        [InlineData(90.0, true)]
        [InlineData(45.0, true)]
        [InlineData(270.0, true)]
        [InlineData(-90.0, true)]
        [InlineData(30.0, false)]
        [InlineData(200.0, false)]
        public void 回転角を180度で折り返して45度から135度の画像を記録する(double rotation, bool expected)
        {
            var sheet = Sheet() with { DrawingObjects = new DrawingObjectModel[] { Image(5, rotation) } };

            var recorded = Compute(sheet).FormatApproximations.Any(a => a.Kind == FormatApproximationKind.ImageRotation);

            Assert.Equal(expected, recorded);
        }

        [Fact]
        public void 反転した画像と同じアンカーの別の画像をそれぞれ記録する()
        {
            var sheet = Sheet() with { DrawingObjects = new DrawingObjectModel[] { Image(5, 0, flip: true), Image(6, 0, flip: true) } };

            var flips = Compute(sheet).FormatApproximations.Where(a => a.Kind == FormatApproximationKind.ImageFlip).ToList();

            Assert.Equal(2, flips.Count);
            Assert.All(flips, f => Assert.Equal(CellRange.Parse("B2:B2"), f.Location));
        }

        [Fact]
        public void グループの中の画像はグループの回転を足した角度と反転で判定する()
        {
            var child = new GroupChildImage(9, new RectPt(0, 0, 10, 10), new byte[] { 1 }, "image/png", 30.0);
            var group = new GroupShapeModel(
                2, new PointPt(0, 0), new PointPt(10, 10), new GroupChildModel[] { child }, 60.0,
                CellAddress.Parse("A1"), new PointPt(0, 0), new FixedAnchorExtent(20.0, 20.0),
                FlipHorizontal: true);
            var sheet = Sheet() with { DrawingObjects = new DrawingObjectModel[] { group } };

            var layout = Compute(sheet);

            Assert.Equal(CellRange.Parse("A1:A1"), Single(layout, FormatApproximationKind.ImageRotation).Location);
            Assert.Equal("rot=90", Single(layout, FormatApproximationKind.ImageRotation).Detail);
            Assert.NotNull(Single(layout, FormatApproximationKind.ImageFlip));
        }

        [Fact]
        public void 印刷範囲の外のセルは記録しない()
        {
            var layout = Compute(Sheet(("E1", Text("範囲外", CellStyle.Default with { HAlign = HorizontalAlignment.Distributed }))));

            Assert.Empty(layout.FormatApproximations);
        }

        [Fact]
        public void 保存では集めず描画命令も変わらない()
        {
            var sheet = Sheet(("A1", Text("承認者", CellStyle.Default with { HAlign = HorizontalAlignment.Distributed })));

            var saving = Compute(sheet, checkFit: false);
            var checking = Compute(sheet, checkFit: true);

            Assert.Empty(saving.FormatApproximations);
            Assert.NotEmpty(checking.FormatApproximations);
            Assert.Equal(
                saving.Pages.SelectMany(p => p.Commands).Count(),
                checking.Pages.SelectMany(p => p.Commands).Count());
        }
    }
}
