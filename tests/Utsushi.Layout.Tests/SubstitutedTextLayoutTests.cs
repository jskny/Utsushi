using System.Collections.Generic;
using System.Globalization;
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
    /// 差し込み文字列の配置(要件2.13, 2.14, 4.6, 4.7)。
    /// </summary>
    public sealed class SubstitutedTextLayoutTests
    {
        private static readonly CellAddress A1 = new(1, 1);

        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

        /// <summary>A1に指定の文字列・書式を持つ1セルのシートを作る。</summary>
        private static SheetModel SingleCellSheet(
            string text, bool wrap, double columnWidth = 6.0, double rowHeightPt = 80.0, VerticalAlignment? vAlign = null)
        {
            var style = CellStyle.Default with { WrapText = wrap };
            if (vAlign is { } align)
            {
                style = style with { VAlign = align };
            }

            var cells = new Dictionary<CellAddress, CellModel>
            {
                [A1] = new(text, CellValueKind.Text, style, text),
            };

            return UniformSheet(rows: 1, columns: 1, columnWidth: columnWidth, rowHeightPt: rowHeightPt,
                pageSetup: NoMarginA4()) with
            { Cells = cells };
        }

        private PagedLayout Compute(SheetModel sheet, params CellAddress[] substituted) =>
            _engine.Compute(ReportModel.Create(Definition(), sheet) with
            {
                SubstitutedCells = new HashSet<CellAddress>(substituted),
            });

        private List<string> WrappedLines(string text, double columnWidth = 6.0) =>
            Texts(Assert.Single(Compute(SingleCellSheet(text, wrap: true, columnWidth)).Pages))
                .Select(t => t.Text)
                .ToList();

        // -- 要件4.6: 書記素クラスタ単位の折り返し ------------------------------

        [Fact]
        public void サロゲートペアの途中で改行しない()
        {
            const string text = "𠮷𠮷𠮷𠮷𠮷𠮷𠮷𠮷𠮷";

            var lines = WrappedLines(text);

            Assert.True(lines.Count > 1, $"折り返されるはず (行数: {lines.Count})");
            Assert.Equal(text, string.Concat(lines));
            Assert.All(lines, line =>
            {
                Assert.False(char.IsLowSurrogate(line[0]), $"行頭が下位サロゲート: {line}");
                Assert.False(char.IsHighSurrogate(line[line.Length - 1]), $"行末が上位サロゲート: {line}");
            });
        }

        [Theory]
        [InlineData("が")] // 合成用濁点
        [InlineData("葛\U000E0100")] // 異体字セレクタ(IVS)
        public void 結合文字や異体字セレクタを基底文字から切り離さない(string cluster)
        {
            var text = string.Concat(Enumerable.Repeat(cluster, 8));

            var lines = WrappedLines(text);

            Assert.True(lines.Count > 1, $"折り返されるはず (行数: {lines.Count})");
            Assert.All(lines, line =>
            {
                var elements = StringInfo.GetTextElementEnumerator(line);
                while (elements.MoveNext())
                {
                    Assert.Equal(cluster, elements.GetTextElement());
                }
            });
        }

        [Theory]
        [InlineData("東京都\r\n丸の内")]
        [InlineData("東京都\r丸の内")]
        [InlineData("東京都\n丸の内")]
        public void 折り返し表示ではCRLFとCR単独も改行として扱う(string text)
        {
            var lines = WrappedLines(text, columnWidth: 20.0);

            Assert.Equal(new[] { "東京都", "丸の内" }, lines);
        }

        // -- 要件4.7: 折り返し表示でないセルの改行 -----------------------------

        [Theory]
        [InlineData("東京都\n丸の内")]
        [InlineData("東京都\r\n丸の内")]
        public void 折り返し表示でないセルでは改行を表示せず1行にする(string text)
        {
            var sheet = SingleCellSheet(text, wrap: false, columnWidth: 20.0, rowHeightPt: 20.0);

            var line = Assert.Single(Texts(Assert.Single(Compute(sheet).Pages)));

            Assert.Equal("東京都丸の内", line.Text);
        }

        // -- 要件2.14: 折り返した差し込み値の欠落 -------------------------------

        [Fact]
        public void 差し込み値の折り返し行がセルの高さに収まらなければエラーになる()
        {
            var sheet = SingleCellSheet("あいうえおかきくけこさしすせそ", wrap: true, rowHeightPt: 20.0);

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet, A1));

            Assert.Equal(A1, ex.CellAddress);
            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Contains("A1", ex.Message);
        }

        [Fact]
        public void テンプレート自身の文字列はセルの高さに収まらなくてもExcelと同様に切り取るだけ()
        {
            var sheet = SingleCellSheet("あいうえおかきくけこさしすせそ", wrap: true, rowHeightPt: 20.0);

            var texts = Texts(Assert.Single(Compute(sheet).Pages)).ToList();

            Assert.True(texts.Count > 1);
            Assert.All(texts, t => Assert.NotNull(t.ClipRect));
        }

        [Fact]
        public void 差し込み値の折り返し行がセルに収まればエラーにならない()
        {
            var sheet = SingleCellSheet("あいうえおかきくけこさしすせそ", wrap: true, rowHeightPt: 120.0);

            var texts = Texts(Assert.Single(Compute(sheet, A1).Pages)).ToList();

            Assert.Equal("あいうえおかきくけこさしすせそ", string.Concat(texts.Select(t => t.Text)));
        }

        [Theory]
        [InlineData(VerticalAlignment.Bottom)]
        [InlineData(VerticalAlignment.Top)]
        [InlineData(VerticalAlignment.Center)]
        public void 行間がセルの高さをわずかに超えるだけならエラーにしない(VerticalAlignment vAlign)
        {
            // 既定フォント11ptの近似行間は13.2pt。2行で26.4ptとなり24ptの行高をわずかに超えるが、
            // どちらの行も半分以上は見えているため欠落とはみなさない。
            var sheet = SingleCellSheet("一行目\n二行目", wrap: true, columnWidth: 20.0, rowHeightPt: 24.0, vAlign: vAlign);

            var texts = Texts(Assert.Single(Compute(sheet, A1).Pages)).ToList();

            Assert.Equal(2, texts.Count);
        }

        [Fact]
        public void 一行目の字面が大きく切れる場合はエラーになる()
        {
            // 1行用の高さ(20pt)に2行を差し込むと、下揃えでは1行目の字面の半分以上がセルの上に出る。
            var sheet = SingleCellSheet("東京都千代田区\n丸の内1-1-1", wrap: true, columnWidth: 20.0, rowHeightPt: 20.0);

            Assert.Throws<LayoutComputationException>(() => Compute(sheet, A1));
        }

        [Fact]
        public void 改ページをまたぐ結合セルでも本来の高さで判定する()
        {
            var text = string.Concat(Enumerable.Repeat("あ", 60));
            var cells = new Dictionary<CellAddress, CellModel>
            {
                [A1] = new(text, CellValueKind.Text, CellStyle.Default with { WrapText = true }, text),
            };
            var sheet = UniformSheet(
                rows: 4, columns: 1, columnWidth: 6.0, rowHeightPt: 20.0,
                pageSetup: NoMarginA4(rowBreaks: new[] { 1 }),
                mergedRanges: new[] { CellRange.Parse("A1:A3") }) with
            { Cells = cells };

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet, A1));

            Assert.Equal(A1, ex.CellAddress);
        }

        [Fact]
        public void 非表示の行にある差し込みセルはエラーになる()
        {
            var sheet = UniformSheet(rows: 3, columns: 1, pageSetup: NoMarginA4()) with
            {
                HiddenRows = new HashSet<int> { 2 },
            };
            var hidden = CellAddress.Parse("A2");

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet, hidden));

            Assert.Equal(hidden, ex.CellAddress);
            Assert.Contains("非表示", ex.Message);
        }

        [Fact]
        public void 幅が無い列にある差し込みセルはエラーになる()
        {
            var sheet = SingleCellSheet("山田 太郎", wrap: false, columnWidth: 0.1, rowHeightPt: 20.0);

            Assert.Throws<LayoutComputationException>(() => Compute(sheet, A1));
        }

        // -- 要件2.13: 差し込み値のセルが印刷範囲外 -----------------------------

        [Fact]
        public void 値を差し込んだ任意セルが印刷範囲外ならエラーになる()
        {
            var sheet = UniformSheet(
                rows: 10, columns: 6,
                pageSetup: NoMarginA4(printAreas: new[] { CellRange.Parse("A1:B2") }));
            var outside = CellAddress.Parse("E8");

            var ex = Assert.Throws<LayoutComputationException>(() => Compute(sheet, outside));

            Assert.Equal(outside, ex.CellAddress);
            Assert.Contains("E8", ex.Message);
        }

        [Fact]
        public void 値を差し込んでいないセルが印刷範囲外でもエラーにならない()
        {
            var sheet = UniformSheet(
                rows: 10, columns: 6,
                pageSetup: NoMarginA4(printAreas: new[] { CellRange.Parse("A1:B2") }));

            var layout = Compute(sheet, CellAddress.Parse("B2"));

            Assert.Single(layout.Pages);
        }
    }
}
