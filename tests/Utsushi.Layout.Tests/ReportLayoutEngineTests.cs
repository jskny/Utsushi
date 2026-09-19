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

namespace Utsushi.Layout.Tests;

/// <summary>
/// レイアウト計算の検証(要件3.1-3.5, 4.1-4.4、タスク5.2-5.8)。
/// </summary>
/// <remarks>
/// フォントメトリクスは実行環境のフォント構成に依存しないよう
/// <see cref="ApproximateFontMetricsProvider"/> を使う。
/// </remarks>
public sealed class ReportLayoutEngineTests
{
    private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

    private PagedLayout Compute(SheetModel sheet, ReportDefinition? definition = null) =>
        _engine.Compute(ReportModel.Create(definition ?? Definition(), sheet));

    // -- 要件3.1: 印刷範囲 -------------------------------------------------

    [Fact]
    public void 印刷範囲の外のセルは出力しない()
    {
        var sheet = UniformSheet(
            rows: 10, columns: 6,
            pageSetup: NoMarginA4(printAreas: new[] { CellRange.Parse("B2:C4") }));

        var layout = Compute(sheet);

        var page = Assert.Single(layout.Pages);
        var texts = Texts(page).Select(t => t.Text).ToHashSet();

        Assert.Equal(new[] { "B2", "B3", "B4", "C2", "C3", "C4" }.ToHashSet(), texts);
    }

    [Fact]
    public void 帳票定義の印刷範囲はExcelの設定より優先される()
    {
        var sheet = UniformSheet(
            rows: 10, columns: 6,
            pageSetup: NoMarginA4(printAreas: new[] { CellRange.Parse("A1:F10") }));

        var layout = Compute(sheet, Definition(printAreaOverride: CellRange.Parse("A1:B1")));

        var texts = Texts(Assert.Single(layout.Pages)).Select(t => t.Text).ToHashSet();
        Assert.Equal(new[] { "A1", "B1" }.ToHashSet(), texts);
    }

    [Fact]
    public void 複数の印刷範囲はそれぞれ別のページ群になる()
    {
        // 要件3.6: 範囲を包含する矩形にまとめず、定義順に独立したページ群として出力する。
        var sheet = UniformSheet(
            rows: 10, columns: 4,
            pageSetup: NoMarginA4(printAreas: new[]
            {
                CellRange.Parse("A1:B2"),
                CellRange.Parse("C8:D9"),
            }));

        var layout = Compute(sheet);

        Assert.Equal(2, layout.PageCount);

        var first = Texts(layout.Pages[0]).Select(t => t.Text).ToHashSet();
        var second = Texts(layout.Pages[1]).Select(t => t.Text).ToHashSet();

        Assert.Equal(new[] { "A1", "A2", "B1", "B2" }.ToHashSet(), first);
        Assert.Equal(new[] { "C8", "C9", "D8", "D9" }.ToHashSet(), second);
    }

    [Fact]
    public void 複数の印刷範囲の間にあるセルは出力されない()
    {
        var sheet = UniformSheet(
            rows: 6, columns: 2,
            pageSetup: NoMarginA4(printAreas: new[]
            {
                CellRange.Parse("A1:B2"),
                CellRange.Parse("A5:B6"),
            }));

        var layout = Compute(sheet);
        var allTexts = layout.Pages.SelectMany(Texts).Select(t => t.Text).ToHashSet();

        // 3〜4行目はどちらの印刷範囲にも含まれない
        Assert.DoesNotContain("A3", allTexts);
        Assert.DoesNotContain("A4", allTexts);
        Assert.Contains("A1", allTexts);
        Assert.Contains("A5", allTexts);
    }

    [Fact]
    public void 複数の印刷範囲でもページ番号は通しで振られる()
    {
        // 1つ目の範囲が2ページに分かれ、2つ目の範囲が1ページになる構成
        var sheet = UniformSheet(
            rows: 30, columns: 1, rowHeightPt: 100.0,
            pageSetup: NoMarginA4(printAreas: new[]
            {
                CellRange.Parse("A1:A20"),
                CellRange.Parse("A25:A26"),
            }));

        var layout = Compute(sheet);

        Assert.Equal(4, layout.PageCount);
        Assert.Equal(new[] { 1, 2, 3, 4 }, layout.Pages.Select(p => p.PageNumber));
    }

    [Fact]
    public void 帳票定義の印刷範囲指定は複数指定より優先される()
    {
        var sheet = UniformSheet(
            rows: 10, columns: 4,
            pageSetup: NoMarginA4(printAreas: new[]
            {
                CellRange.Parse("A1:B2"),
                CellRange.Parse("C8:D9"),
            }));

        var layout = Compute(sheet, Definition(printAreaOverride: CellRange.Parse("A1:A1")));

        var page = Assert.Single(layout.Pages);
        Assert.Equal("A1", Assert.Single(Texts(page)).Text);
    }

    [Fact]
    public void 印刷範囲もセルも無ければエラーになる()
    {
        var sheet = UniformSheet(rows: 0, columns: 0, pageSetup: NoMarginA4());

        var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet));

        Assert.Equal(ProcessingStage.Layout, ex.Stage);
    }

    // -- 要件3.2/3.3: 改ページ -------------------------------------------

    [Fact]
    public void 行高の合計が印字可能領域を超えたら改ページする()
    {
        // A4縦・余白0の高さ = 841.89pt。行高100pt → 8行/ページ
        var sheet = UniformSheet(rows: 20, columns: 1, rowHeightPt: 100.0, pageSetup: NoMarginA4());

        var layout = Compute(sheet);

        Assert.Equal(3, layout.PageCount);
        Assert.Equal((1, 8), layout.Pages[0].RowRange);
        Assert.Equal((9, 16), layout.Pages[1].RowRange);
        Assert.Equal((17, 20), layout.Pages[2].RowRange);
    }

    [Fact]
    public void 手動改ページの位置で分割する()
    {
        var sheet = UniformSheet(
            rows: 10, columns: 1, rowHeightPt: 10.0,
            pageSetup: NoMarginA4(rowBreaks: new[] { 5 }));

        var layout = Compute(sheet);

        Assert.Equal(2, layout.PageCount);
        Assert.Equal((1, 4), layout.Pages[0].RowRange);
        Assert.Equal((5, 10), layout.Pages[1].RowRange);
    }

    [Fact]
    public void 列方向にも改ページする()
    {
        // A4縦・余白0の幅 = 595.28pt。列幅 10文字 = 70px = 52.5pt → 11列/ページ
        var sheet = UniformSheet(rows: 1, columns: 25, columnWidth: 10.0, pageSetup: NoMarginA4());

        var layout = Compute(sheet);

        Assert.Equal(3, layout.PageCount);
        Assert.Equal((1, 11), layout.Pages[0].ColumnRange);
        Assert.Equal((12, 22), layout.Pages[1].ColumnRange);
        Assert.Equal((23, 25), layout.Pages[2].ColumnRange);
    }

    // -- 要件3.4: 印刷タイトル -------------------------------------------

    [Fact]
    public void 印刷タイトル行は各ページの先頭に繰り返される()
    {
        var sheet = UniformSheet(
            rows: 20, columns: 1, rowHeightPt: 100.0,
            pageSetup: NoMarginA4(printTitles: new PrintTitles(1, 2, null, null)));

        var layout = Compute(sheet);

        // タイトル2行(200pt)を除いた残り641.89pt → 本文6行/ページ
        Assert.Equal(3, layout.PageCount);

        foreach (var page in layout.Pages)
        {
            var texts = Texts(page).Select(t => t.Text).ToList();
            Assert.Contains("A1", texts);
            Assert.Contains("A2", texts);
        }

        // 本文はタイトル行を含まず、3行目から流し込まれる
        Assert.Equal((3, 8), layout.Pages[0].RowRange);
        Assert.Equal((9, 14), layout.Pages[1].RowRange);
    }

    [Fact]
    public void 印刷タイトルは2ページ目以降でページ先頭に配置される()
    {
        var sheet = UniformSheet(
            rows: 20, columns: 1, rowHeightPt: 100.0,
            pageSetup: NoMarginA4(printTitles: new PrintTitles(1, 1, null, null)));

        var layout = Compute(sheet);
        var secondPage = layout.Pages[1];

        var titleText = Texts(secondPage).Single(t => t.Text == "A1");
        var firstBodyText = Texts(secondPage).Single(t => t.Text == "A" + secondPage.RowRange.First);

        Assert.True(
            titleText.Origin.Y < firstBodyText.Origin.Y,
            "タイトル行は本文より上に配置されるはず");
    }

    // -- 要件3.5: 印刷順序 -----------------------------------------------

    [Fact]
    public void 既定の印刷順序は上から下そのあと右へ()
    {
        // 2列帯 × 2行帯 になるサイズ
        var sheet = UniformSheet(rows: 20, columns: 25, rowHeightPt: 100.0, columnWidth: 10.0,
            pageSetup: NoMarginA4());

        var layout = Compute(sheet);

        // 先に行方向(下)へ進み、行帯を使い切ってから次の列帯へ移る
        Assert.Equal((1, 8), layout.Pages[0].RowRange);
        Assert.Equal((1, 11), layout.Pages[0].ColumnRange);
        Assert.Equal((9, 16), layout.Pages[1].RowRange);
        Assert.Equal((1, 11), layout.Pages[1].ColumnRange);
    }

    [Fact]
    public void 印刷順序が列優先なら左から右へ進む()
    {
        var sheet = UniformSheet(rows: 20, columns: 25, rowHeightPt: 100.0, columnWidth: 10.0,
            pageSetup: NoMarginA4(pageOrder: PageOrder.OverThenDown));

        var layout = Compute(sheet);

        Assert.Equal((1, 8), layout.Pages[0].RowRange);
        Assert.Equal((1, 11), layout.Pages[0].ColumnRange);
        Assert.Equal((1, 8), layout.Pages[1].RowRange);
        Assert.Equal((12, 22), layout.Pages[1].ColumnRange);
    }

    // -- 拡大縮小 ---------------------------------------------------------

    [Fact]
    public void 縮小率を指定すると1ページに入る行数が増える()
    {
        var sheet = UniformSheet(rows: 20, columns: 1, rowHeightPt: 100.0,
            pageSetup: NoMarginA4(scaling: new PageScaling(50, null, null)));

        var layout = Compute(sheet);

        // 50%縮小 → 論理的な印字可能高さが倍の1683.78pt → 16行/ページ
        Assert.Equal(2, layout.PageCount);
        Assert.Equal((1, 16), layout.Pages[0].RowRange);
        Assert.Equal(0.5, layout.Pages[0].ScaleFactor);
    }

    [Fact]
    public void 指定ページ数に収める設定なら1ページに収まる()
    {
        var sheet = UniformSheet(rows: 20, columns: 1, rowHeightPt: 100.0,
            pageSetup: NoMarginA4(scaling: new PageScaling(100, null, 1)));

        var layout = Compute(sheet);

        Assert.Equal(1, layout.PageCount);
        Assert.True(layout.Pages[0].ScaleFactor < 1.0);
    }

    [Fact]
    public void 余白が用紙を超えるとエラーになる()
    {
        var sheet = UniformSheet(rows: 1, columns: 1);
        var pageSetup = PageSetupModel.Default with
        {
            Margins = new PageMargins(400, 400, 500, 500, 0, 0),
        };

        var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet with { PageSetup = pageSetup }));

        Assert.Equal(ProcessingStage.Layout, ex.Stage);
    }

    // -- 要件4.2/4.3: 罫線と結合セル ------------------------------------

    [Fact]
    public void 結合セルは1つの矩形として扱われる()
    {
        var style = CellStyle.Default with { BackgroundColor = new ArgbColor(0xFF, 0xFF, 0x00, 0x00) };
        var sheet = UniformSheet(
            rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0,
            pageSetup: NoMarginA4(),
            mergedRanges: new[] { CellRange.Parse("A1:C1") },
            style: style);

        var layout = Compute(sheet);
        var page = Assert.Single(layout.Pages);

        // 結合範囲は左上セルの値だけを1回描画する
        var texts = Texts(page).Select(t => t.Text).ToList();
        Assert.Contains("A1", texts);
        Assert.DoesNotContain("B1", texts);
        Assert.DoesNotContain("C1", texts);

        // 背景は3列ぶんの幅を持つ1つの矩形になる
        var columnWidthPt = ExcelUnitConverter.ColumnWidthToPoints(10.0, 7.0);
        var mergedFill = Fills(page).Single(f => f.Rect.Top == 0.0 && f.Rect.Left == 0.0);
        Assert.Equal(columnWidthPt * 3, mergedFill.Rect.Width, precision: 6);
    }

    [Fact]
    public void 罫線は上下左右の各辺として出力される()
    {
        var border = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
        var style = CellStyle.Default with
        {
            Borders = new BorderSet(border, border, border, border, BorderEdge.None, BorderEdge.None),
        };

        var sheet = UniformSheet(rows: 1, columns: 1, pageSetup: NoMarginA4(), style: style);

        var lines = Lines(Assert.Single(Compute(sheet).Pages)).ToList();

        Assert.Equal(4, lines.Count);
        Assert.All(lines, l => Assert.Equal(Units.PixelsToPoints(1.0), l.WidthPt, precision: 6));
    }

    [Fact]
    public void 隣接セルの同じ罫線は重複して出力されない()
    {
        var border = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
        var style = CellStyle.Default with
        {
            Borders = new BorderSet(border, border, border, border, BorderEdge.None, BorderEdge.None),
        };

        // 横に2セル並べると、間の境界は「左セルの右辺」と「右セルの左辺」で重なる
        var sheet = UniformSheet(rows: 1, columns: 2, pageSetup: NoMarginA4(), style: style);

        var lines = Lines(Assert.Single(Compute(sheet).Pages)).ToList();

        // 縦4本ではなく3本(重複が除かれる)
        var verticalCount = lines.Count(l => l.From.X == l.To.X);
        Assert.Equal(3, verticalCount);
    }

    // -- 要件4.1/4.4/2.5: テキスト配置 -----------------------------------

    [Theory]
    [InlineData(HorizontalAlignment.Left, TextAnchor.Left)]
    [InlineData(HorizontalAlignment.Center, TextAnchor.Center)]
    [InlineData(HorizontalAlignment.Right, TextAnchor.Right)]
    public void 水平配置が描画命令に反映される(HorizontalAlignment align, TextAnchor expected)
    {
        var sheet = UniformSheet(
            rows: 1, columns: 1, pageSetup: NoMarginA4(),
            style: CellStyle.Default with { HAlign = align });

        var text = Assert.Single(Texts(Assert.Single(Compute(sheet).Pages)));

        Assert.Equal(expected, text.Anchor);
    }

    [Fact]
    public void 標準配置では文字列は左寄せ数値は右寄せになる()
    {
        var cells = new Dictionary<CellAddress, CellModel>
        {
            [new CellAddress(1, 1)] = new("文字", CellValueKind.Text, CellStyle.Default, "文字"),
            [new CellAddress(2, 1)] = new("100", CellValueKind.Number, CellStyle.Default, "100"),
        };

        var sheet = UniformSheet(rows: 2, columns: 1, pageSetup: NoMarginA4()) with { Cells = cells };

        var texts = Texts(Assert.Single(Compute(sheet).Pages)).ToList();

        Assert.Equal(TextAnchor.Left, texts.Single(t => t.Text == "文字").Anchor);
        Assert.Equal(TextAnchor.Right, texts.Single(t => t.Text == "100").Anchor);
    }

    [Theory]
    [InlineData(VerticalAlignment.Top)]
    [InlineData(VerticalAlignment.Center)]
    [InlineData(VerticalAlignment.Bottom)]
    public void 垂直配置でベースライン位置が変わる(VerticalAlignment align)
    {
        var sheet = UniformSheet(
            rows: 1, columns: 1, rowHeightPt: 60.0, pageSetup: NoMarginA4(),
            style: CellStyle.Default with { VAlign = align });

        var text = Assert.Single(Texts(Assert.Single(Compute(sheet).Pages)));
        var metrics = new ApproximateFontMetricsProvider().GetMetrics(FontStyle.Default);

        var expected = align switch
        {
            VerticalAlignment.Top => metrics.AscentPt,
            VerticalAlignment.Center => ((60.0 - metrics.LineSpacingPt) / 2.0) + metrics.AscentPt,
            _ => 60.0 - metrics.LineSpacingPt + metrics.AscentPt,
        };

        Assert.Equal(expected, text.Origin.Y, precision: 6);
    }

    [Fact]
    public void 縮小指定ならセル幅に収まるまでフォントを縮める()
    {
        const string longText = "株式会社ながいなまえの取引先さま";
        var cells = new Dictionary<CellAddress, CellModel>
        {
            [new CellAddress(1, 1)] = new(longText, CellValueKind.Text, CellStyle.Default, longText),
        };

        var sheet = UniformSheet(rows: 1, columns: 1, columnWidth: 4.0, pageSetup: NoMarginA4()) with
        {
            Cells = cells,
        };

        var definition = Definition(
            fields: new[]
            {
                new SubstitutionFieldDefinition(
                    "Long", CellAddress.Parse("A1"), false, OverflowBehavior.Shrink),
            });

        var report = ReportModel.Create(definition, sheet) with
        {
            OverflowByCell = new Dictionary<CellAddress, OverflowBehavior>
            {
                [CellAddress.Parse("A1")] = OverflowBehavior.Shrink,
            },
        };

        var text = Assert.Single(Texts(Assert.Single(_engine.Compute(report).Pages)));

        Assert.True(
            text.Font.SizePt < FontStyle.Default.SizePt,
            $"縮小されるはず (実際: {text.Font.SizePt}pt)");
        Assert.NotNull(text.ClipRect);

        // 縮小後の文字列幅がセルの内容領域に収まっていること
        var metrics = new ApproximateFontMetricsProvider();
        var contentWidth = ExcelUnitConverter.ColumnWidthToPoints(4.0, 7.0)
            - (ExcelUnitConverter.CellPaddingPoints * 2);
        Assert.True(
            metrics.MeasureTextWidth(text.Font, text.Text) <= contentWidth + 0.001,
            "縮小後はセル幅に収まるはず");
    }

    [Fact]
    public void 折り返し指定なら複数行に分割される()
    {
        var cells = new Dictionary<CellAddress, CellModel>
        {
            [new CellAddress(1, 1)] = new(
                "あいうえおかきくけこさしすせそ", CellValueKind.Text,
                CellStyle.Default with { WrapText = true },
                "あいうえおかきくけこさしすせそ"),
        };

        var sheet = UniformSheet(rows: 1, columns: 1, columnWidth: 6.0, rowHeightPt: 80.0,
            pageSetup: NoMarginA4()) with
        { Cells = cells };

        var texts = Texts(Assert.Single(Compute(sheet).Pages)).ToList();

        Assert.True(texts.Count > 1, $"折り返されるはず (行数: {texts.Count})");
        Assert.Equal("あいうえおかきくけこさしすせそ", string.Concat(texts.Select(t => t.Text)));
    }

    [Fact]
    public void はみ出し許容ならクリップしない()
    {
        var sheet = UniformSheet(rows: 1, columns: 1, columnWidth: 2.0, pageSetup: NoMarginA4());

        var text = Assert.Single(Texts(Assert.Single(Compute(sheet).Pages)));

        Assert.Null(text.ClipRect);
    }

    [Fact]
    public void ページ座標には余白が加算される()
    {
        var sheet = UniformSheet(
            rows: 1, columns: 1,
            pageSetup: PageSetupModel.Default with
            {
                Margins = new PageMargins(50.0, 10.0, 30.0, 10.0, 0, 0),
            });

        var page = Assert.Single(Compute(sheet).Pages);

        // 罫線が無いので、テキストの基準位置で余白の適用を確認する
        var text = Assert.Single(Texts(page));
        Assert.True(text.Origin.X >= 50.0, $"左余白が反映されるはず (X={text.Origin.X})");
        Assert.True(text.Origin.Y >= 30.0, $"上余白が反映されるはず (Y={text.Origin.Y})");
    }

    [Fact]
    public void 描画順は背景_罫線_テキストになる()
    {
        var border = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
        var style = CellStyle.Default with
        {
            BackgroundColor = new ArgbColor(0xFF, 0xEE, 0xEE, 0xEE),
            Borders = new BorderSet(border, border, border, border, BorderEdge.None, BorderEdge.None),
        };

        var sheet = UniformSheet(rows: 2, columns: 2, pageSetup: NoMarginA4(), style: style);
        var commands = Assert.Single(Compute(sheet).Pages).Commands;

        var lastFill = commands.Select((c, i) => (c, i)).Last(x => x.c is FillRectCommand).i;
        var firstLine = commands.Select((c, i) => (c, i)).First(x => x.c is LineCommand).i;
        var lastLine = commands.Select((c, i) => (c, i)).Last(x => x.c is LineCommand).i;
        var firstText = commands.Select((c, i) => (c, i)).First(x => x.c is TextCommand).i;

        Assert.True(lastFill < firstLine, "背景はすべて罫線より前");
        Assert.True(lastLine < firstText, "罫線はすべてテキストより前");
    }

    [Fact]
    public void 非表示の行と列は出力しない()
    {
        var sheet = UniformSheet(rows: 3, columns: 3, pageSetup: NoMarginA4()) with
        {
            HiddenRows = new HashSet<int> { 2 },
            HiddenColumns = new HashSet<int> { 2 },
        };

        var texts = Texts(Assert.Single(Compute(sheet).Pages)).Select(t => t.Text).ToHashSet();

        Assert.Equal(new[] { "A1", "C1", "A3", "C3" }.ToHashSet(), texts);
    }
}
