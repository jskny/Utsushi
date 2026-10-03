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
    /// 文字の収まりの確認(要件13)と、セルが無い位置の書式(要件1.10)・数値書式の色(要件4.12)の描画。
    /// </summary>
    /// <remarks>
    /// 近似メトリクス(半角0.5em・全角1em)で、11ptの半角30文字は約165pt。列幅10文字(最大数字幅7px)の列は約56ptなので、
    /// 1列に収まらず、隣の2列にかかる。
    /// </remarks>
    public sealed class FitCheckAndEffectiveStyleLayoutTests
    {
        private const string LongText = "ABCDEFGHIJABCDEFGHIJABCDEFGHIJ";

        private static readonly ArgbColor Red = new(0xFF, 0xFF, 0x00, 0x00);

        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

        /// <summary>3行×5列(列幅10文字・行高20pt)の、指定セルだけを持つシート。</summary>
        private static SheetModel Sheet(PageSetupModel? pageSetup = null, params (string Cell, CellModel Model)[] cells)
        {
            var map = cells.ToDictionary(c => CellAddress.Parse(c.Cell), c => c.Model);
            return UniformSheet(rows: 3, columns: 5, pageSetup: pageSetup ?? NoMarginA4(
                printAreas: new[] { CellRange.Parse("A1:E3") })) with
            { Cells = map };
        }

        private static CellModel Text(string text, CellStyle? style = null) =>
            new(text, CellValueKind.Text, style ?? CellStyle.Default, text);

        private static CellModel Number(string formatted, CellStyle? style = null) =>
            new("0", CellValueKind.Number, style ?? CellStyle.Default, formatted);

        private PagedLayout Compute(SheetModel sheet, ReportDefinition? definition = null, params CellAddress[] substituted) =>
            _engine.Compute(ReportModel.Create(definition ?? Definition(), sheet) with
            {
                SubstitutedCells = new HashSet<CellAddress>(substituted),
            });

        // --- 要件13: 文字の収まりの確認 ---------------------------------------------

        [Fact]
        public void はみ出した文字が値のある隣のセルに重なれば検出する()
        {
            var layout = Compute(Sheet(null, ("A1", Text(LongText)), ("C1", Text("値"))));

            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(FitIssueKind.OverlapsNeighborValue, issue.Kind);
            Assert.Equal(CellAddress.Parse("A1"), issue.Cell);
            Assert.Equal(1, issue.PageNumber);
            Assert.Equal(LongText, issue.Text);
            Assert.False(issue.IsSubstituted);
            Assert.Contains("C1", issue.Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void はみ出した先が空のセルなら検出しない()
        {
            var layout = Compute(Sheet(null, ("A1", Text(LongText)), ("E1", Text("遠くの値"))));

            Assert.Empty(layout.FitIssues);
        }

        [Fact]
        public void 右揃えの文字は左側の隣のセルとの重なりを調べる()
        {
            var right = CellStyle.Default with { HAlign = HorizontalAlignment.Right };
            var layout = Compute(Sheet(null, ("B1", Text("ラベル")), ("D1", Text(LongText, right))));

            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(FitIssueKind.OverlapsNeighborValue, issue.Kind);
            Assert.Equal(CellAddress.Parse("D1"), issue.Cell);
            Assert.Contains("B1", issue.Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void 隣の結合範囲にアンカーの値があれば重なりとして検出する()
        {
            var sheet = Sheet(null, ("A1", Text(LongText)), ("B1", Text("結合")))
                with
            { MergedRanges = new[] { new MergedRange(CellRange.Parse("B1:C1")) } };

            var issue = Assert.Single(Compute(sheet).FitIssues);
            Assert.Equal(FitIssueKind.OverlapsNeighborValue, issue.Kind);
            Assert.Contains("B1", issue.Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void ページの右端ではみ出した文字が切れれば検出する()
        {
            var layout = Compute(Sheet(null, ("E1", Text(LongText))));

            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(FitIssueKind.CutAtPageEdge, issue.Kind);
            Assert.Equal(CellAddress.Parse("E1"), issue.Cell);
        }

        [Fact]
        public void 数値がセルの幅に収まらなければ検出し描画は変えない()
        {
            var layout = Compute(Sheet(null, ("A1", Number("1,234,567,890,123", CellStyle.Default with { NumberFormat = "#,##0" }))));

            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(FitIssueKind.NumberTooWide, issue.Kind);
            Assert.Contains("####", issue.Message, System.StringComparison.Ordinal);

            // #### にはせず、数値の文字列のまま描く(要件13.2)。
            Assert.Contains(Texts(Assert.Single(layout.Pages)), t => t.Text == "1,234,567,890,123");
        }

        [Fact]
        public void 切り取り表示と結合セルの文字が切れれば検出する()
        {
            var clip = Definition(fields: new[]
            {
                new SubstitutionFieldDefinition("Name", CellAddress.Parse("A1"), false, OverflowBehavior.Clip),
            });
            var sheet = Sheet(null, ("A1", Text(LongText)), ("A3", Text(LongText)))
                with
            { MergedRanges = new[] { new MergedRange(CellRange.Parse("A3:B3")) } };

            var issues = _engine.Compute(ReportModel.Create(clip, sheet) with
            {
                OverflowByCell = new Dictionary<CellAddress, OverflowBehavior> { [CellAddress.Parse("A1")] = OverflowBehavior.Clip },
                SubstitutedCells = new HashSet<CellAddress> { CellAddress.Parse("A1") },
            }).FitIssues;

            Assert.Equal(2, issues.Count);
            Assert.All(issues, i => Assert.Equal(FitIssueKind.Clipped, i.Kind));

            var substituted = issues.Single(i => i.Cell == CellAddress.Parse("A1"));
            Assert.True(substituted.IsSubstituted);
            Assert.Equal("Name", substituted.SubstitutionKey);
            Assert.Contains("置換キー 'Name'", substituted.Message, System.StringComparison.Ordinal);

            Assert.Contains("A3:B3", issues.Single(i => i.Cell == CellAddress.Parse("A3")).Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void テンプレートの折り返した文字がセルの高さに収まらなければ検出する()
        {
            var wrap = CellStyle.Default with { WrapText = true };
            var layout = Compute(Sheet(null, ("A1", Text(LongText, wrap)), ("A2", Text("短い", wrap))));

            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(FitIssueKind.ExceedsCellHeight, issue.Kind);
            Assert.Equal(CellAddress.Parse("A1"), issue.Cell);
        }

        [Fact]
        public void 縮小表示のセルは検出しない()
        {
            var shrink = CellStyle.Default with { ShrinkToFit = true };
            var layout = Compute(Sheet(null, ("A1", Text(LongText, shrink)), ("B1", Text("値")), ("C1", Number("123456789012", shrink))));

            Assert.Empty(layout.FitIssues);
        }

        [Fact]
        public void 印刷タイトルとして複数ページに現れるセルは最初のページの1件にまとめる()
        {
            var pageSetup = NoMarginA4(
                printAreas: new[] { CellRange.Parse("A1:E3") },
                rowBreaks: new[] { 2, 3 },
                printTitles: new PrintTitles(1, 1, null, null));

            var layout = Compute(Sheet(pageSetup, ("A1", Text(LongText)), ("B1", Text("値"))));

            Assert.True(layout.PageCount >= 2);
            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(1, issue.PageNumber);
        }

        [Fact]
        public void 中央揃えの文字は左右両方の隣のセルとの重なりを調べる()
        {
            var center = CellStyle.Default with { HAlign = HorizontalAlignment.Center };
            var layout = Compute(Sheet(null, ("A1", Text("左")), ("C1", Text(LongText + LongText, center))));

            var issue = Assert.Single(layout.FitIssues, i => i.Kind == FitIssueKind.OverlapsNeighborValue);
            Assert.Contains("A1", issue.Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void 数式の結果が空文字の隣のセルも値があるものとして扱う()
        {
            // =IF(…,"",…) の結果が空文字でも、Excel ははみ出した文字をその手前で止める。
            var formula = new CellModel(string.Empty, CellValueKind.Text, CellStyle.Default, string.Empty, HasFormula: true);
            var layout = Compute(Sheet(null, ("A1", Text(LongText)), ("B1", formula)));

            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(FitIssueKind.OverlapsNeighborValue, issue.Kind);
        }

        [Fact]
        public void 説明の文にはセルの文字列を含めない()
        {
            var layout = Compute(Sheet(null, ("A1", Text("株式会社サンプル商事 東日本統括本部 御中")), ("B1", Text("値"))));

            var issue = Assert.Single(layout.FitIssues);
            Assert.DoesNotContain("サンプル", issue.Message, System.StringComparison.Ordinal);
            Assert.Equal("株式会社サンプル商事 東日本統括本部 御中", issue.Text);
        }

        [Fact]
        public void 標準の表示形式の数値はExcelで桁を減らして表示される旨を説明する()
        {
            var general = Number("1234567890123456");
            var withFormat = Number("1,234,567,890,123", CellStyle.Default with { NumberFormat = "#,##0" });

            var issues = Compute(Sheet(null, ("A1", general), ("A2", withFormat))).FitIssues;

            Assert.Contains("桁を減らす", issues.Single(i => i.Cell == CellAddress.Parse("A1")).Message, System.StringComparison.Ordinal);
            Assert.Contains("####", issues.Single(i => i.Cell == CellAddress.Parse("A2")).Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void 改ページをまたぐ結合範囲は範囲全体の幅で判定し見えている部分の幅で誤検出しない()
        {
            // A1:B1 の結合範囲を列の手動改ページ(A列の後)で分ける。文字は2列ぶんなら収まる。
            var pageSetup = NoMarginA4(printAreas: new[] { CellRange.Parse("A1:E3") }, columnBreaks: new[] { 2 });
            var sheet = Sheet(pageSetup, ("A1", Text("ABCDEFGHIJABCDEF")))
                with
            { MergedRanges = new[] { new MergedRange(CellRange.Parse("A1:B1")) } };

            var layout = Compute(sheet);

            Assert.True(layout.PageCount >= 2);
            Assert.Empty(layout.FitIssues);
        }

        [Fact]
        public void 折り返しの高さは描画と同じ縦位置で判定する()
        {
            // 行高20pt・近似メトリクス(字面 1.1em、行送り 1.1em 程度)で2行の文字。上下中央なら上下に少しずつはみ出すだけで、
            // 差し込み値の判定(要件2.14)と同じく、字面の4分の1以内なら収まるとみなす。
            var centered = CellStyle.Default with { WrapText = true, VAlign = VerticalAlignment.Center, Font = FontStyle.Default with { SizePt = 9 } };
            var layout = Compute(Sheet(null, ("A1", Text("ABCDEFGHIJABCDEFGH", centered))));

            Assert.Empty(layout.FitIssues);
        }

        [Fact]
        public void セル番地直接指定で上書きしたセルには置換キーを付けない()
        {
            var definition = Definition(fields: new[]
            {
                new SubstitutionFieldDefinition("Name", CellAddress.Parse("A1"), false, null),
            });
            var a1 = CellAddress.Parse("A1");

            var layout = _engine.Compute(ReportModel.Create(definition, Sheet(null, ("A1", Text(LongText)), ("B1", Text("値")))) with
            {
                SubstitutedCells = new HashSet<CellAddress> { a1 },
                OverriddenCells = new HashSet<CellAddress> { a1 },
            });

            var issue = Assert.Single(layout.FitIssues);
            Assert.True(issue.IsSubstituted);
            Assert.Null(issue.SubstitutionKey);
            Assert.DoesNotContain("置換キー", issue.Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void 置換していないセルの説明の文には置換キーを付けない()
        {
            var definition = Definition(fields: new[]
            {
                new SubstitutionFieldDefinition("Name", CellAddress.Parse("A1"), false, null),
            });

            var issue = Assert.Single(Compute(Sheet(null, ("A1", Text(LongText)), ("B1", Text("値"))), definition).FitIssues);

            Assert.False(issue.IsSubstituted);
            Assert.DoesNotContain("置換キー", issue.Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void 収まりの確認を省くとFitIssuesは空になる()
        {
            var report = ReportModel.Create(Definition(), Sheet(null, ("A1", Text(LongText)), ("B1", Text("値"))));

            Assert.Empty(_engine.Compute(report, checkFit: false).FitIssues);
            Assert.Single(_engine.Compute(report, checkFit: true).FitIssues);
        }

        [Fact]
        public void 件数の上限に達したら打ち切りを示す()
        {
            var collector = new FitIssueCollector();
            for (var i = 0; i < FitIssueCollector.MaxIssues; i++)
            {
                collector.Add(new FitIssue(FitIssueKind.Clipped, new CellAddress(i + 1, 1), 1, "x", false, null, "m"));
            }

            Assert.False(collector.IsTruncated);
            Assert.False(collector.Accepts(new CellAddress(1, 2), FitIssueKind.Clipped));
            Assert.True(collector.IsTruncated);
            Assert.Equal(FitIssueCollector.MaxIssues, collector.Issues.Count);
        }

        [Fact]
        public void 折り返し表示のセルの数値は折り返さずセルで切り取り収まらなければ検出する()
        {
            // Excel は数値を折り返さず、隣のセルへはみ出させもしない(要件14.3)。
            var wrap = CellStyle.Default with { WrapText = true };
            var layout = Compute(Sheet(null, ("A1", Text("見出し")), ("B1", Number("1234567890123456789", wrap))));

            var page = Assert.Single(layout.Pages);
            var number = Assert.Single(Texts(page), t => t.Text == "1234567890123456789");
            Assert.NotNull(number.ClipRect);

            // B1 のセルの矩形(列幅10文字、約56pt)で切り取り、左隣の A1 の見出しへは描かない。
            Assert.True(number.ClipRect!.Value.Width < 60.0, "セルの矩形で切り取るはず");
            Assert.True(number.ClipRect.Value.Left > 50.0, "A1 の上には描かないはず");

            var issue = Assert.Single(layout.FitIssues);
            Assert.Equal(FitIssueKind.NumberTooWide, issue.Kind);
        }

        [Fact]
        public void 折り返し表示のセルの文字列は従来どおり折り返す()
        {
            var wrap = CellStyle.Default with { WrapText = true };
            var page = Assert.Single(Compute(Sheet(null, ("A1", Text("ABCDEFGHIJABCDEF", wrap)))).Pages);

            Assert.True(Texts(page).Count() >= 2);
        }

        // --- 要件1.10: セルが無い位置の書式 -----------------------------------------

        [Fact]
        public void セルが無い位置にも行の書式の塗りつぶしと罫線を描く()
        {
            var yellow = new ArgbColor(0xFF, 0xFF, 0xFF, 0x00);
            var thin = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
            var band = CellStyle.Default with
            {
                BackgroundColor = yellow,
                Borders = new BorderSet(BorderEdge.None, BorderEdge.None, BorderEdge.None, thin, BorderEdge.None, BorderEdge.None),
            };
            var sheet = Sheet(null, ("A1", Text("見出し"))) with
            {
                RowStyles = new Dictionary<int, CellStyle> { [2] = band },
            };

            var page = Assert.Single(Compute(sheet).Pages);

            // 同じ行に並ぶ5セルぶんの塗りつぶしは、1つの矩形にまとめる。
            var fill = Assert.Single(Fills(page), f => f.Color == yellow);
            Assert.Equal(20.0, fill.Rect.Top, 6);
            Assert.True(fill.Rect.Width > 4 * 50.0);
            Assert.NotEmpty(Lines(page));
        }

        [Fact]
        public void 列の書式は行の書式より優先度が低く印刷範囲の外には描かない()
        {
            var yellow = new ArgbColor(0xFF, 0xFF, 0xFF, 0x00);
            var blue = new ArgbColor(0xFF, 0x00, 0x00, 0xFF);
            var sheet = Sheet(NoMarginA4(printAreas: new[] { CellRange.Parse("A1:C3") }), ("A1", Text("見出し"))) with
            {
                RowStyles = new Dictionary<int, CellStyle> { [2] = CellStyle.Default with { BackgroundColor = yellow } },
                ColumnStyles = new[] { new ColumnStyleRange(2, 4, CellStyle.Default with { BackgroundColor = blue }) },
            };

            var page = Assert.Single(Compute(sheet).Pages);
            var fills = Fills(page).ToList();

            // 行2は行の書式(黄)が A〜C 列に1つ、列の書式(青)は行1・行3の B〜C 列だけ(D列は印刷範囲の外)。
            Assert.Single(fills, f => f.Color == yellow);
            Assert.Equal(2, fills.Count(f => f.Color == blue));
            Assert.All(fills.Where(f => f.Color == blue), f => Assert.True(f.Rect.Right <= fills.Single(y => y.Color == yellow).Rect.Right + 1e-6));
        }

        [Fact]
        public void 行列の書式で描くセルの無い位置が上限を超えたら中止する()
        {
            var counter = new StyledBlankPositionCounter("test-report", "テストシート", limit: 3);
            counter.Count();
            counter.Count();
            counter.Count();

            Assert.Throws<Utsushi.Core.Exceptions.LayoutComputationException>(() => counter.Count());
        }

        [Fact]
        public void 列全体の書式で印刷範囲の全位置を描かせる入力は上限で止まる()
        {
            var fill = CellStyle.Default with { BackgroundColor = new ArgbColor(0xFF, 0xFF, 0xFF, 0xFF) };
            var sheet = UniformSheet(rows: 1, columns: 1, pageSetup: NoMarginA4(
                printAreas: new[] { new CellRange(1, 1, 20_000, 20) })) with
            {
                ColumnStyles = new[] { new ColumnStyleRange(1, CellAddress.MaxColumn, fill) },
            };

            var ex = Assert.Throws<Utsushi.Core.Exceptions.LayoutComputationException>(
                () => _engine.Compute(ReportModel.Create(Definition(), sheet)));
            Assert.Contains("上限", ex.Message, System.StringComparison.Ordinal);
        }

        [Fact]
        public void 結合範囲の外周の罫線はセルが無い位置の書式からも探す()
        {
            var thin = new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black);
            var bottom = CellStyle.Default with
            {
                Borders = new BorderSet(BorderEdge.None, BorderEdge.None, BorderEdge.None, thin, BorderEdge.None, BorderEdge.None),
            };
            var sheet = Sheet(null, ("A1", Text("結合"))) with
            {
                MergedRanges = new[] { new MergedRange(CellRange.Parse("A1:A2")) },
                RowStyles = new Dictionary<int, CellStyle> { [2] = bottom },
            };

            var page = Assert.Single(Compute(sheet).Pages);

            // A1:A2 の下辺(行2の下端 = 40pt)に罫線が出る。
            Assert.Contains(Lines(page), l => l.From.Y == 40.0 && l.To.Y == 40.0);
        }

        // --- 要件4.12: 数値書式の色 ----------------------------------------------------

        [Fact]
        public void 数値書式の色があればその色で数値を描く()
        {
            var negative = Number("-3,000") with { FormatColor = Red };
            var page = Assert.Single(Compute(Sheet(null, ("A1", negative), ("A2", Number("3,000")))).Pages);

            Assert.Equal(Red, Texts(page).Single(t => t.Text == "-3,000").Font.Color);
            Assert.Equal(ArgbColor.Black, Texts(page).Single(t => t.Text == "3,000").Font.Color);
        }
    }
}
