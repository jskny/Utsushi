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
            var layout = Compute(Sheet(null, ("A1", Number("1,234,567,890,123"))));

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

            Assert.Equal(5, Fills(page).Count(f => f.Color == yellow));
            Assert.NotEmpty(Lines(page));
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
