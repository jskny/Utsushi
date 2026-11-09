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
            Assert.Contains(horizontal, a => a.Location == CellRange.Parse("A1:A1") && a.Detail == "horizontal=distributed");
            Assert.Contains(horizontal, a => a.Location == CellRange.Parse("B1:B1") && a.Detail == "horizontal=centerContinuous");
            Assert.Equal("vertical=justify", Single(layout, FormatApproximationKind.VerticalAlignment).Detail);
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
                Vertical = "eaVert",
            };
            var shape = new ShapeModel(
                7u, ShapePresetType.Rect, System.Array.Empty<double>(), 0, null, null, text,
                CellAddress.Parse("B2"), new PointPt(0, 0), new FixedAnchorExtent(40.0, 30.0));
            var sheet = Sheet() with { DrawingObjects = new DrawingObjectModel[] { shape } };

            var rotation = Single(Compute(sheet), FormatApproximationKind.TextRotation);

            Assert.Equal(CellRange.Parse("B2:B2"), rotation.Location);
            Assert.Equal("vert=eaVert", rotation.Detail);
        }

        [Fact]
        public void 行の書式の網かけはセルが無い位置も印刷される範囲だけ記録する()
        {
            var patterned = CellStyle.Default with
            {
                BackgroundColor = new ArgbColor(0xFF, 0xEE, 0xEE, 0xEE),
                FillApproximation = FormatApproximationKind.CellFillPattern,
                FillApproximationDetail = "patternType=darkTrellis",
            };
            var sheet = Sheet() with { RowStyles = new Dictionary<int, CellStyle> { [1] = patterned } };

            var fills = Compute(sheet).FormatApproximations.Where(a => a.Kind == FormatApproximationKind.CellFillPattern).ToList();

            // 印刷範囲 A1:C3 の1行目だけ(D1・E1 は印刷範囲の外)。
            Assert.Equal(new[] { "A1:A1", "B1:B1", "C1:C1" }, fills.Select(f => f.Location!.Value.ToString()).OrderBy(x => x).ToArray());
            Assert.All(fills, f => Assert.Equal("patternType=darkTrellis", f.Detail));
        }

        private static ImageModel Image(uint id, double rotation, bool flip = false) =>
            new(id, new byte[] { 1 }, "image/png", rotation, CellAddress.Parse("B2"), new PointPt(0, 0), new FixedAnchorExtent(20.0, 10.0))
            {
                FlipHorizontal = flip,
            };

        [Theory]
        [InlineData(90.0, true)]
        [InlineData(45.0, true)]
        [InlineData(134.9, true)]
        [InlineData(135.0, false)]
        [InlineData(270.0, true)]
        [InlineData(315.0, false)]
        [InlineData(-90.0, true)]
        [InlineData(30.0, false)]
        [InlineData(200.0, false)]
        public void 回転角を180度で折り返して45度以上135度未満の画像を記録する(double rotation, bool expected)
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
        public void グループの中の画像は回転を記録せず外側のグループの反転を含めて反転を記録する()
        {
            var child = new GroupChildImage(9, new RectPt(0, 0, 10, 10), new byte[] { 1 }, "image/png", 30.0);
            var group = new GroupShapeModel(
                2, new PointPt(0, 0), new PointPt(10, 10), new GroupChildModel[] { child }, 60.0,
                CellAddress.Parse("A1"), new PointPt(0, 0), new FixedAnchorExtent(20.0, 20.0),
                FlipHorizontal: true);
            var sheet = Sheet() with { DrawingObjects = new DrawingObjectModel[] { group } };

            var layout = Compute(sheet);

            // 子の位置・大きさは回転前の矩形で、中心で回して描くため、回転は近似ではない(layout-fidelity-reviewer指摘)。
            Assert.DoesNotContain(layout.FormatApproximations, a => a.Kind == FormatApproximationKind.ImageRotation);
            Assert.Equal(CellRange.Parse("A1:A1"), Single(layout, FormatApproximationKind.ImageFlip).Location);
        }

        [Fact]
        public void 複数行のヘッダーとフッターを位置ごとに1件だけ記録する()
        {
            var headerFooter = new HeaderFooterModel(
                OddHeader: "&L1行目\n2行目&R右", OddFooter: "&Cページ\n&P",
                EvenHeader: null, EvenFooter: null,
                FirstHeader: null, FirstFooter: null,
                DifferentOddEven: false, DifferentFirst: false, ScaleWithDocument: true);
            var sheet = UniformSheet(rows: 120, columns: 3, pageSetup: NoMarginA4() with { HeaderFooter = headerFooter });

            var layout = Compute(sheet);

            Assert.True(layout.Pages.Count > 1);
            var items = layout.FormatApproximations.Where(a => a.Kind == FormatApproximationKind.MultiLineHeaderFooter).ToList();
            Assert.Equal(new[] { "footer/center", "header/left" }, items.Select(i => i.Detail).OrderBy(d => d).ToArray());
            Assert.All(items, i => Assert.Null(i.Location));
        }

        [Fact]
        public void 先頭ページ用と偶数ページ用のヘッダーは識別に別を含める()
        {
            var headerFooter = new HeaderFooterModel(
                OddHeader: "&L奇数", OddFooter: null,
                EvenHeader: "&L偶数\n2行目", EvenFooter: null,
                FirstHeader: "&R先頭\n2行目", FirstFooter: null,
                DifferentOddEven: true, DifferentFirst: true, ScaleWithDocument: true);
            var sheet = UniformSheet(rows: 120, columns: 3, pageSetup: NoMarginA4() with { HeaderFooter = headerFooter });

            var layout = Compute(sheet);

            Assert.True(layout.Pages.Count > 2);
            var items = layout.FormatApproximations.Where(a => a.Kind == FormatApproximationKind.MultiLineHeaderFooter).ToList();
            Assert.Equal(new[] { "evenHeader/left", "firstHeader/right" }, items.Select(i => i.Detail).OrderBy(d => d).ToArray());
        }

        private static SheetModel LongSheet(params CellRange[] merged) =>
            UniformSheet(rows: 120, columns: 3, pageSetup: NoMarginA4(), mergedRanges: merged);

        private static ShapeModel WideShape(string anchor, double widthPt) =>
            new(
                3u, ShapePresetType.Rect, System.Array.Empty<double>(), 0, null, null, null,
                CellAddress.Parse(anchor), new PointPt(0, 0), new FixedAnchorExtent(widthPt, 10.0));

        private static ShapeModel TallShape(string anchor, double heightPt) =>
            new(
                3u, ShapePresetType.Rect, System.Array.Empty<double>(), 0, null, null, null,
                CellAddress.Parse(anchor), new PointPt(0, 0), new FixedAnchorExtent(40.0, heightPt));

        [Fact]
        public void 改ページをまたぐ結合セルを1件だけ記録する()
        {
            var layout = Compute(LongSheet(CellRange.Parse("A40:A45")));

            Assert.True(layout.Pages.Count > 1);
            Assert.Equal(CellRange.Parse("A40:A45"), Single(layout, FormatApproximationKind.MergedCellAcrossPageBreak).Location);
        }

        [Fact]
        public void 一つのページに収まる結合セルは記録しない()
        {
            var layout = Compute(LongSheet(CellRange.Parse("A2:B3")));

            Assert.DoesNotContain(layout.FormatApproximations, a => a.Kind == FormatApproximationKind.MergedCellAcrossPageBreak);
        }

        [Fact]
        public void 最後の行が非表示で下端に罫線のある結合セルを記録する()
        {
            var thin = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
            var bottomBorder = CellStyle.Default with
            {
                Borders = new BorderSet(BorderEdge.None, BorderEdge.None, BorderEdge.None, thin, BorderEdge.None, BorderEdge.None),
            };
            var sheet = Sheet(("A3", Text(string.Empty, bottomBorder))) with
            {
                HiddenRows = new HashSet<int> { 3 },
                MergedRanges = new[] { new MergedRange(CellRange.Parse("A2:B3")) },
            };

            var item = Single(Compute(sheet), FormatApproximationKind.MergedCellHiddenEdge);

            Assert.Equal(CellRange.Parse("A2:B3"), item.Location);
            Assert.Equal("lastRowHidden", item.Detail);
        }

        [Fact]
        public void 最後の列が非表示で右端に罫線のある結合セルを記録する()
        {
            var thin = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
            var rightBorder = CellStyle.Default with
            {
                Borders = new BorderSet(BorderEdge.None, thin, BorderEdge.None, BorderEdge.None, BorderEdge.None, BorderEdge.None),
            };
            var sheet = Sheet(("B2", Text(string.Empty, rightBorder))) with
            {
                HiddenColumns = new HashSet<int> { 2 },
                MergedRanges = new[] { new MergedRange(CellRange.Parse("A2:B2")) },
            };

            Assert.Equal("lastColumnHidden", Single(Compute(sheet), FormatApproximationKind.MergedCellHiddenEdge).Detail);
        }

        [Fact]
        public void 最後の行と列が両方非表示の結合セルは両方を補足の情報に含める()
        {
            var thin = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
            var borders = CellStyle.Default with
            {
                Borders = new BorderSet(BorderEdge.None, thin, BorderEdge.None, thin, BorderEdge.None, BorderEdge.None),
            };
            var sheet = Sheet(("B3", Text(string.Empty, borders))) with
            {
                HiddenRows = new HashSet<int> { 3 },
                HiddenColumns = new HashSet<int> { 2 },
                MergedRanges = new[] { new MergedRange(CellRange.Parse("A2:B3")) },
            };

            Assert.Equal(
                "lastRowHidden,lastColumnHidden", Single(Compute(sheet), FormatApproximationKind.MergedCellHiddenEdge).Detail);
        }

        [Fact]
        public void 右の続きのページへはみ出す図形を記録する()
        {
            // 20列は A4 縦の1ページの幅に収まらず、右に続きのページがある。
            var sheet = UniformSheet(rows: 3, columns: 20, pageSetup: NoMarginA4()) with
            {
                DrawingObjects = new DrawingObjectModel[] { WideShape("C2", 1000.0) },
            };

            var item = Single(Compute(sheet), FormatApproximationKind.DrawingAcrossPageBreak);

            Assert.Equal(CellRange.Parse("C2:C2"), item.Location);
        }

        [Fact]
        public void 最後の行が非表示でも罫線の無い結合セルは記録しない()
        {
            // 描かない辺に罫線が無ければ、見た目は Excel と変わらない(layout-fidelity-reviewer指摘)。
            var sheet = Sheet() with
            {
                HiddenRows = new HashSet<int> { 3 },
                MergedRanges = new[] { new MergedRange(CellRange.Parse("A2:B3")) },
            };

            Assert.DoesNotContain(Compute(sheet).FormatApproximations, a => a.Kind == FormatApproximationKind.MergedCellHiddenEdge);
        }

        [Fact]
        public void 印刷範囲の末尾の列が非表示なら右へのはみ出しを改ページをまたぐ図形として記録しない()
        {
            // 印刷範囲 A1:E3 のうち D:E が非表示で、1ページに収まる。続きのページは無い(layout-fidelity-reviewer指摘)。
            var sheet = UniformSheet(rows: 3, columns: 5, pageSetup: NoMarginA4(printAreas: new[] { CellRange.Parse("A1:E3") })) with
            {
                HiddenColumns = new HashSet<int> { 4, 5 },
                DrawingObjects = new DrawingObjectModel[] { WideShape("C2", 300.0) },
            };

            var layout = Compute(sheet);

            Assert.Equal(1, layout.Pages.Count);
            Assert.DoesNotContain(layout.FormatApproximations, a => a.Kind == FormatApproximationKind.DrawingAcrossPageBreak);
        }

        [Fact]
        public void 印刷タイトルにかかる条件付き書式を記録する()
        {
            // 印刷タイトルの行(1〜4行目)は印刷範囲の外にあるが、各ページに印刷される(要件7.3)。
            var sheet = UniformSheet(
                rows: 120,
                columns: 3,
                pageSetup: NoMarginA4(
                    printAreas: new[] { CellRange.Parse("A5:C120") },
                    printTitles: new PrintTitles(1, 4, null, null))) with
            {
                ConditionalFormattingRanges = new[] { CellRange.Parse("A1:C4") },
            };

            Assert.Equal(CellRange.Parse("A1:C4"), Single(Compute(sheet), FormatApproximationKind.ConditionalFormatting).Location);
        }

        [Fact]
        public void 条件付き書式は印刷される行の範囲だけを記録し非表示の行だけにかかるものは記録しない()
        {
            var sheet = Sheet() with
            {
                HiddenRows = new HashSet<int> { 3 },
                ConditionalFormattingRanges = new[] { CellRange.Parse("A2:C3"), CellRange.Parse("A3:B3") },
            };

            var item = Single(Compute(sheet), FormatApproximationKind.ConditionalFormatting);

            Assert.Equal(CellRange.Parse("A2:C2"), item.Location);
        }

        [Fact]
        public void 改ページをまたぐ図形を記録する()
        {
            var sheet = LongSheet() with { DrawingObjects = new DrawingObjectModel[] { TallShape("A40", 200.0) } };

            Assert.Equal(CellRange.Parse("A40:A40"), Single(Compute(sheet), FormatApproximationKind.DrawingAcrossPageBreak).Location);
        }

        [Fact]
        public void 続きのページが無い方向へのはみ出しは改ページをまたぐ図形として記録しない()
        {
            var sheet = Sheet() with { DrawingObjects = new DrawingObjectModel[] { TallShape("A2", 500.0) } };

            Assert.DoesNotContain(Compute(sheet).FormatApproximations, a => a.Kind == FormatApproximationKind.DrawingAcrossPageBreak);
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
