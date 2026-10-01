using System;
using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.HeaderFooter;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using static Utsushi.Layout.Tests.LayoutFixtures;

namespace Utsushi.Layout.Tests
{
    /// <summary>
    /// 監査で見つかった Layout レイヤーの不具合(印刷タイトルの位置、ページ数に合わせる、結合セル・はみ出し表示のクリップ、
    /// 枠線の倍率、幅0・高さ0の行列、改ページ位置、二重罫線、ページ中央、開始ページ番号、図形内テキストの余白・改行)の回帰テスト。
    /// </summary>
    public sealed class LayoutFidelityRegressionTests
    {
        /// <summary>A4縦の高さ(pt)。余白ゼロのときの印字可能領域の高さ。</summary>
        private const double A4HeightPt = 841.8897637795275;

        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider(), () => new DateTime(2026, 4, 20));

        private PagedLayout Compute(SheetModel sheet, params CellAddress[] substituted) =>
            _engine.Compute(ReportModel.Create(Definition(), sheet) with
            {
                SubstitutedCells = new HashSet<CellAddress>(substituted),
            });

        private static List<string> TextsOf(PageLayout page) => Texts(page).Select(t => t.Text).ToList();

        // -- 印刷タイトルの位置 ----------------------------------------------------

        [Fact]
        public void 印刷範囲の中にあるタイトル行は1ページ目の行順を崩さず2ページ目以降にだけ付く()
        {
            // 印刷範囲 1〜120、タイトル $3:$4。1ページ目は A1,A2,A3,... の順のまま。
            var sheet = UniformSheet(
                rows: 120, columns: 1, rowHeightPt: 100.0,
                pageSetup: NoMarginA4(
                    printAreas: new[] { CellRange.Parse("A1:A120") },
                    printTitles: new PrintTitles(3, 4, null, null)));

            var layout = Compute(sheet);

            Assert.Equal(Enumerable.Range(1, 8).Select(r => "A" + r), TextsOf(layout.Pages[0]));
            Assert.Equal((1, 8), layout.Pages[0].RowRange);

            // 2ページ目は帯の先頭(9行目)がタイトルの最終行(4行目)より後ろなので、タイトルを付ける。
            // タイトル2行(200pt)の分だけ本文は6行に減る。
            Assert.Equal(
                new[] { "A3", "A4" }.Concat(Enumerable.Range(9, 6).Select(r => "A" + r)),
                TextsOf(layout.Pages[1]));
            Assert.Equal((9, 14), layout.Pages[1].RowRange);
        }

        [Fact]
        public void タイトル行の手前で始まるページにはタイトルを付けない()
        {
            // タイトル $10:$11。1ページ目(1〜8行)と2ページ目(9行目から)はタイトルより前から始まる。
            var sheet = UniformSheet(
                rows: 30, columns: 1, rowHeightPt: 100.0,
                pageSetup: NoMarginA4(printTitles: new PrintTitles(10, 11, null, null)));

            var layout = Compute(sheet);

            Assert.Equal(Enumerable.Range(9, 8).Select(r => "A" + r), TextsOf(layout.Pages[1]));
            Assert.Equal(new[] { "A10", "A11", "A17" }, TextsOf(layout.Pages[2]).Take(3));
        }

        [Fact]
        public void 印刷範囲の中にあるタイトル列は1ページ目の列順を崩さず2ページ目以降にだけ付く()
        {
            var sheet = UniformSheet(
                rows: 1, columns: 40, columnWidth: 10.0,
                pageSetup: NoMarginA4(printTitles: new PrintTitles(null, null, 2, 3)));

            var layout = Compute(sheet);
            Assert.True(layout.PageCount >= 2);

            var first = TextsOf(layout.Pages[0]);
            Assert.Equal(new[] { "A1", "B1", "C1", "D1" }, first.Take(4));
            Assert.Equal(1, layout.Pages[0].ColumnRange.First);

            var second = TextsOf(layout.Pages[1]);
            var secondBodyFirst = new CellAddress(1, layout.Pages[1].ColumnRange.First).ToString();
            Assert.Equal(new[] { "B1", "C1", secondBodyFirst }, second.Take(3));
        }

        [Fact]
        public void どのページにも付かない印刷範囲より下のタイトル行へ差し込むとエラーになる()
        {
            // タイトル $10:$10 は印刷範囲 A1:A5 より下にあり、タイトルより後ろから始まるページが無い。
            var sheet = UniformSheet(
                rows: 10, columns: 1,
                pageSetup: NoMarginA4(
                    printAreas: new[] { CellRange.Parse("A1:A5") },
                    printTitles: new PrintTitles(10, 10, null, null)));

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet, CellAddress.Parse("A10")));
            Assert.Contains("印刷範囲の外", ex.Message);
        }

        // -- ページ数に合わせる ------------------------------------------------------

        [Fact]
        public void ちょうど収まる倍率では浮動小数点の誤差で次のページへ送らない()
        {
            // 合計が印字可能領域のちょうど2倍 → 50%で1ページに収まる。
            var sheet = UniformSheet(
                rows: 10, columns: 1, rowHeightPt: A4HeightPt / 5.0,
                pageSetup: NoMarginA4(scaling: new PageScaling(100, null, 1)));

            var layout = Compute(sheet);

            var page = Assert.Single(layout.Pages);
            Assert.Equal(0.5, page.ScaleFactor, 10);
        }

        [Fact]
        public void 行の区切り位置のせいで指定ページ数を超える場合は倍率を整数パーセントずつ下げる()
        {
            // 0.6ページ分の行が3行。合計(1.8ページ)は2ページに収まるが、100%では1行ずつ3ページになる。
            // 2行が1ページに入るのは 1/1.2 = 83.3% 以下 → 83%。
            var sheet = UniformSheet(
                rows: 3, columns: 1, rowHeightPt: A4HeightPt * 0.6,
                pageSetup: NoMarginA4(scaling: new PageScaling(100, null, 2)));

            var layout = Compute(sheet);

            Assert.Equal(2, layout.PageCount);
            Assert.Equal(0.83, layout.Pages[0].ScaleFactor, 10);
        }

        [Fact]
        public void ページ数に合わせる倍率は10パーセントより下げない()
        {
            var sheet = UniformSheet(
                rows: 1, columns: 1, rowHeightPt: A4HeightPt * 100.0,
                pageSetup: NoMarginA4(scaling: new PageScaling(100, null, 1)));

            var layout = Compute(sheet);

            Assert.Equal(0.1, Assert.Single(layout.Pages).ScaleFactor, 10);
        }

        [Fact]
        public void ページ数に合わせるときは手動改ページを無視する()
        {
            var sheet = UniformSheet(
                rows: 5, columns: 1,
                pageSetup: NoMarginA4(rowBreaks: new[] { 3 }, scaling: new PageScaling(100, null, 1)));

            var page = Assert.Single(Compute(sheet).Pages);
            Assert.Equal((1, 5), page.RowRange);
            Assert.Equal(1.0, page.ScaleFactor);
        }

        // -- 結合セル・はみ出し表示のクリップ ----------------------------------------

        [Fact]
        public void 結合セルのはみ出し表示の文字は結合範囲でクリップする()
        {
            var sheet = UniformSheet(
                rows: 1, columns: 4, columnWidth: 2.0, pageSetup: NoMarginA4(),
                mergedRanges: new[] { CellRange.Parse("A1:B1") });

            var page = Assert.Single(Compute(sheet).Pages);
            var text = Texts(page).Single(t => t.Text == "A1");

            var columnWidthPt = ExcelUnitConverter.ColumnWidthToPoints(2.0, ReportDefinition.DefaultMaxDigitWidthPx);
            Assert.NotNull(text.ClipRect);
            Assert.Equal(0.0, text.ClipRect!.Value.Left, 6);
            Assert.Equal(columnWidthPt * 2, text.ClipRect.Value.Right, 6);
        }

        [Fact]
        public void はみ出し表示の文字は余白の外へ描かない()
        {
            // 1列だけで用紙幅を超える列。はみ出し表示の文字は右余白の内側で切り取る。
            var sheet = UniformSheet(
                rows: 1, columns: 1, columnWidth: 150.0,
                pageSetup: PageSetupModel.Default with { Margins = new PageMargins(50, 40, 30, 20, 0, 0) });

            var page = Assert.Single(Compute(sheet).Pages);
            var clip = Assert.Single(Texts(page)).ClipRect;

            Assert.NotNull(clip);
            Assert.Equal(50.0, clip!.Value.Left, 6);
            Assert.Equal(30.0, clip.Value.Top, 6);
            Assert.Equal(page.WidthPt - 40.0, clip.Value.Right, 6);
        }

        // -- 枠線の太さ ------------------------------------------------------------

        [Fact]
        public void 図形と接続線の枠線の太さに印刷倍率が掛かり既定の線も補われる()
        {
            var arrow = new LineEndStyle(LineEndType.Triangle, LineEndSize.Medium, LineEndSize.Medium);
            var shape = new ShapeModel(
                1u, ShapePresetType.Rect, Array.Empty<double>(), 0, null,
                new ShapeOutline(ArgbColor.Black, 2.0, arrow, null), null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(20, 20));
            var noLine = shape with { Id = 2u, Outline = new ShapeOutline(ArgbColor.Transparent, 0.0) };
            var connector = new ConnectorModel(
                ConnectorPresetType.Straight, 0, false, false, null, null, null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(20, 20));
            var sheet = UniformSheet(
                rows: 2, columns: 2, pageSetup: NoMarginA4(scaling: new PageScaling(50, null, null))) with
            {
                DrawingObjects = new DrawingObjectModel[] { shape, noLine, connector },
            };

            var page = Assert.Single(Compute(sheet).Pages);
            var shapes = Shapes(page).ToList();

            Assert.Equal(1.0, shapes[0].Outline!.WidthPt, 10);
            Assert.Equal(arrow, shapes[0].Outline!.HeadEnd);
            Assert.True(shapes[1].Outline!.Color.IsTransparent);

            var connectorOutline = Assert.Single(Connectors(page)).Outline;
            Assert.NotNull(connectorOutline);
            Assert.Equal(ArgbColor.Black, connectorOutline!.Color);
            Assert.Equal(0.5, connectorOutline.WidthPt, 10);
        }

        // -- 幅0・高さ0の行と列 ----------------------------------------------------

        [Fact]
        public void 高さ0の行に差し込んだ値はエラーになる()
        {
            var sheet = UniformSheet(rows: 3, columns: 1, pageSetup: NoMarginA4());
            sheet = sheet with { RowHeights = new List<double> { 20.0, 0.0, 20.0 } };

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet, CellAddress.Parse("A2")));
            Assert.Contains("非表示", ex.Message);
        }

        [Fact]
        public void 幅0の列に差し込んだ値はエラーになる()
        {
            var sheet = UniformSheet(rows: 1, columns: 3, pageSetup: NoMarginA4());
            sheet = sheet with { ColumnWidths = new List<double> { 10.0, 0.0, 10.0 } };

            Assert.Throws<LayoutComputationException>(() => Compute(sheet, CellAddress.Parse("B1")));
        }

        // -- 改ページの位置 -------------------------------------------------------

        [Fact]
        public void 改ページ位置の行が非表示でも改ページする()
        {
            var sheet = UniformSheet(rows: 5, columns: 1, pageSetup: NoMarginA4(rowBreaks: new[] { 3 })) with
            {
                HiddenRows = new HashSet<int> { 3 },
            };

            var layout = Compute(sheet);

            Assert.Equal(2, layout.PageCount);
            Assert.Equal((1, 2), layout.Pages[0].RowRange);
            Assert.Equal((4, 5), layout.Pages[1].RowRange);
        }

        // -- 二重罫線 -------------------------------------------------------------

        [Fact]
        public void 二重罫線の2本の線の中心は2ピクセル離れる()
        {
            var edge = new BorderEdge(BorderLineStyle.Double, ArgbColor.Black);
            var style = CellStyle.Default with
            {
                Borders = new BorderSet(BorderEdge.None, BorderEdge.None, edge, BorderEdge.None, BorderEdge.None, BorderEdge.None),
            };
            var sheet = UniformSheet(rows: 1, columns: 1, pageSetup: NoMarginA4(), style: style);

            var lines = Lines(Assert.Single(Compute(sheet).Pages)).OrderBy(l => l.From.Y).ToList();

            Assert.Equal(2, lines.Count);
            Assert.Equal(1.5, lines[1].From.Y - lines[0].From.Y, 10);
            Assert.Equal(0.75, lines[0].WidthPt, 10);
        }

        // -- ページ中央 -------------------------------------------------------------

        [Fact]
        public void ページ中央を指定すると本文を印字可能領域の中央へ移す()
        {
            var style = CellStyle.Default with { BackgroundColor = new ArgbColor(0xFF, 0xEE, 0xEE, 0xEE) };
            var pageSetup = PageSetupModel.Default with
            {
                Margins = new PageMargins(50, 30, 40, 20, 0, 0),
                HorizontalCentered = true,
                VerticalCentered = true,
            };
            var sheet = UniformSheet(rows: 1, columns: 1, rowHeightPt: 20.0, pageSetup: pageSetup, style: style);

            var page = Assert.Single(Compute(sheet).Pages);
            var fill = Assert.Single(Fills(page)).Rect;

            var printableWidth = page.WidthPt - 50 - 30;
            var printableHeight = page.HeightPt - 40 - 20;
            Assert.Equal(50 + ((printableWidth - fill.Width) / 2.0), fill.Left, 6);
            Assert.Equal(40 + ((printableHeight - 20.0) / 2.0), fill.Top, 6);
        }

        // -- 開始ページ番号 ---------------------------------------------------------

        [Fact]
        public void 開始ページ番号はヘッダーフッターのページ番号と総ページ数に反映される()
        {
            var headerFooter = HeaderFooterModel.None with
            {
                OddFooter = "&C&P/&N",
                FirstFooter = "&C&P/&N",
                FirstHeader = "&C先頭",
                DifferentFirst = true,
            };
            var sheet = UniformSheet(
                rows: 2, columns: 1,
                pageSetup: NoMarginA4(rowBreaks: new[] { 2 }) with { HeaderFooter = headerFooter, FirstPageNumber = 5 });

            var layout = Compute(sheet);

            Assert.Contains("5/6", TextsOf(layout.Pages[0]));
            Assert.Contains("6/6", TextsOf(layout.Pages[1]));

            // 「先頭ページのみ別指定」は、番号によらず文書の先頭ページに付く。
            Assert.Contains("先頭", TextsOf(layout.Pages[0]));
            Assert.DoesNotContain("先頭", TextsOf(layout.Pages[1]));
        }

        // -- 図形内テキストの余白・改行 -------------------------------------------------

        private static readonly FontStyle ShapeFont = new("Calibri", 10.0, false, false, UnderlineStyle.None, false, ArgbColor.Black);

        private static ShapeTextBody TextBody(ShapeTextInsets? insets = null, params ShapeTextParagraph[] paragraphs) =>
            new(paragraphs, VerticalAlignment.Top, insets);

        private static ShapeTextParagraph Paragraph(params ShapeTextRun[] runs) => new(runs, HorizontalAlignment.Left);

        private ShapeCommand LayoutShape(ShapeTextBody text, PageScaling? scaling = null)
        {
            var shape = new ShapeModel(
                1u, ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, text,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(100, 100));
            var sheet = UniformSheet(rows: 2, columns: 2, rowHeightPt: 60.0, pageSetup: NoMarginA4(scaling: scaling)) with
            {
                DrawingObjects = new[] { shape },
            };

            return Assert.Single(Shapes(Assert.Single(Compute(sheet).Pages)));
        }

        [Fact]
        public void 図形内テキストの余白は既定でDrawingMLの既定値になる()
        {
            var command = LayoutShape(TextBody(null, Paragraph(new ShapeTextRun("A", ShapeFont))));

            var line = Assert.Single(command.TextLines);
            var metrics = new ApproximateFontMetricsProvider().GetMetrics(ShapeFont);
            Assert.Equal(7.2, line.Origin.X, 6);
            Assert.Equal(3.6 + metrics.AscentPt, line.Origin.Y, 6);
        }

        [Fact]
        public void 図形内テキストの余白の指定に印刷倍率が掛かる()
        {
            var insets = new ShapeTextInsets(10, 20, 10, 20);
            var command = LayoutShape(
                TextBody(insets, Paragraph(new ShapeTextRun("A", ShapeFont))), new PageScaling(50, null, null));

            var line = Assert.Single(command.TextLines);
            Assert.Equal(5.0, line.Origin.X, 6);
        }

        [Fact]
        public void グループ内の図形の余白にはグループの倍率も掛かる()
        {
            var child = new GroupChildShape(
                1u, RectPt.FromBounds(0, 0, 80, 80), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null,
                TextBody(null, Paragraph(new ShapeTextRun("A", ShapeFont))));
            var group = new GroupShapeModel(
                2u, new PointPt(0, 0), new PointPt(80, 80), new GroupChildModel[] { child }, 0,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(40, 40));
            var sheet = UniformSheet(rows: 2, columns: 2, rowHeightPt: 60.0, pageSetup: NoMarginA4()) with
            {
                DrawingObjects = new[] { group },
            };

            var groupCommand = Assert.Single(Groups(Assert.Single(Compute(sheet).Pages)));
            var shape = Assert.IsType<ShapeCommand>(Assert.Single(groupCommand.Children));

            Assert.Equal(3.6, Assert.Single(shape.TextLines).Origin.X, 6);
        }

        [Fact]
        public void 図形内テキストの段落内改行と空の段落は行になる()
        {
            var large = ShapeFont with { SizePt = 20.0 };
            var command = LayoutShape(
                TextBody(
                    null,
                    Paragraph(new ShapeTextRun("AB\nCD", ShapeFont)),
                    Paragraph(),
                    Paragraph(new ShapeTextRun("EF", large))),
                new PageScaling(50, null, null));

            Assert.Equal(new[] { "AB", "CD", string.Empty, "EF" }, command.TextLines.Select(l => l.Text));

            // 空の段落の行の高さは直前の段落のフォントで決め、印刷倍率も掛ける。
            Assert.Equal(5.0, command.TextLines[2].Font.SizePt, 10);
            Assert.Equal(10.0, command.TextLines[3].Font.SizePt, 10);
        }

        // -- 単位換算 ---------------------------------------------------------------

        [Fact]
        public void NaNの列幅は0ピクセルになる()
        {
            Assert.Equal(0.0, ExcelUnitConverter.ColumnWidthToPixels(double.NaN, 7.0));
        }
    }

    /// <summary>ヘッダー/フッターの書式コード(ページ番号の加減算・日本語のスタイル名)の回帰テスト。</summary>
    public sealed class HeaderFooterCodeRegressionTests
    {
        private static HeaderFooterContext Context(int pageNumber) =>
            new(pageNumber, 9, "シート", "report", new DateTime(2026, 4, 20), FontStyle.Default);

        private static string CenterText(string definition, int pageNumber) =>
            string.Concat(HeaderFooterParser.Parse(definition, Context(pageNumber))
                .Single(p => p.Section == HeaderFooterSection.Center).Runs.Select(r => r.Text));

        [Theory]
        [InlineData("&P+1", 2, "3")]
        [InlineData("&P-1", 2, "1")]
        [InlineData("&P+10ページ", 2, "12ページ")]
        [InlineData("&P+", 2, "2+")]
        [InlineData("&P-x", 2, "2-x")]
        public void ページ番号の加減算を計算する(string definition, int pageNumber, string expected)
        {
            Assert.Equal(expected, CenterText(definition, pageNumber));
        }

        [Theory]
        [InlineData("太字", true, false)]
        [InlineData("斜体", false, true)]
        [InlineData("太字 斜体", true, true)]
        [InlineData("標準", false, false)]
        [InlineData("Bold Italic", true, true)]
        public void 日本語のスタイル名も太字と斜体として扱う(string style, bool bold, bool italic)
        {
            var parts = HeaderFooterParser.Parse($"&C&\"ＭＳ ゴシック,{style}\"文字", Context(1));

            var run = Assert.Single(Assert.Single(parts).Runs);
            Assert.Equal("ＭＳ ゴシック", run.Font.Name);
            Assert.Equal(bold, run.Font.Bold);
            Assert.Equal(italic, run.Font.Italic);
        }
    }

    /// <summary>セル→結合範囲の索引(<see cref="MergedCellIndex"/>)の検証。</summary>
    public sealed class MergedCellIndexTests
    {
        [Fact]
        public void 索引の検索結果はFindMergedRangeと一致する()
        {
            var ranges = new[] { "B2:D4", "F1:F10", "A6:C6", "H3:J3", "C3:E5" }
                .Select(r => new MergedRange(CellRange.Parse(r)))
                .ToList();
            var sheet = UniformSheet(rows: 1, columns: 1) with { MergedRanges = ranges };
            var rows = Enumerable.Range(1, 12).ToList();

            var index = MergedCellIndex.Create(ranges, rows);

            foreach (var row in rows)
            {
                for (var column = 1; column <= 12; column++)
                {
                    var address = new CellAddress(row, column);
                    Assert.Equal(sheet.FindMergedRange(address), index.Find(address));
                }
            }
        }

        [Fact]
        public void 索引に登録していない行は結合なしとして扱う()
        {
            var ranges = new[] { new MergedRange(CellRange.Parse("A1:A100")) };

            var index = MergedCellIndex.Create(ranges, new[] { 1, 50 });

            Assert.NotNull(index.Find(new CellAddress(50, 1)));
            Assert.Null(index.Find(new CellAddress(2, 1)));
        }
    }
}
