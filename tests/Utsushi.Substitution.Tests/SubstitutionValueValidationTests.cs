using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Xunit;
using static Utsushi.Substitution.Tests.TestReports;

namespace Utsushi.Substitution.Tests
{
    /// <summary>
    /// 差し込み値そのものの検証・正規化と、差し込み済みセルの記録(要件2.10〜2.13)。
    /// </summary>
    public sealed class SubstitutionValueValidationTests
    {
        private readonly CellSubstitutor _substitutor = new();

        [Theory]
        [InlineData("INV\t0001", "U+0009")]
        [InlineData("株式会社\u0000サンプル", "U+0000")]
        [InlineData("縦タブ\u000Bあり", "U+000B")]
        [InlineData("DEL\u007F", "U+007F")]
        [InlineData("山田\u200B太郎", "U+200B")] // ゼロ幅スペース
        [InlineData("\uFEFF株式会社", "U+FEFF")] // BOM
        public void 改行以外の制御文字や不可視の書式文字を含む値はエラーになる(string value, string expectedCodePoint)
        {
            var report = Report(Settings(), Sheet(("C3", "x", null)));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["C3"] = value }));

            Assert.Equal("C3", ex.Target);
            Assert.Equal(CellAddress.Parse("C3"), ex.CellAddress);
            Assert.Equal(ProcessingStage.Substitution, ex.Stage);
            Assert.Contains(expectedCodePoint, ex.Message);
        }

        [Theory]
        [InlineData(new[] { 0xD842 })] // 上位サロゲートのみ
        [InlineData(new[] { 'a', 0xDFB7, 'b' })] // 下位サロゲートのみ
        public void 対になっていないサロゲートを含む値はエラーになる(int[] chars)
        {
            // 孤立サロゲートを文字列リテラルのままInlineDataに渡すと、テストデータの直列化で
            // U+FFFDに置き換わってしまうため、UTF-16の単位の配列から組み立てる。
            var value = new string(System.Array.ConvertAll(chars, c => (char)c));
            var report = Report(Settings(), Sheet(("A1", "x", null)));

            Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = value }));
        }

        [Fact]
        public void サロゲートペアの文字は受け付ける()
        {
            var report = Report(Settings(), Sheet(("A1", "x", null)));

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = "𠮷野 様" });

            Assert.Equal("𠮷野 様", result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void nullの値はエラーになる()
        {
            var report = Report(Settings(), Sheet(("A1", "x", null)));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = null! }));

            Assert.Equal("A1", ex.Target);
        }

        [Theory]
        [InlineData("東京都千代田区\r\n丸の内1-1-1")]
        [InlineData("東京都千代田区\r丸の内1-1-1")]
        [InlineData("東京都千代田区\n丸の内1-1-1")]
        public void 改行コードはLFに統一される(string value)
        {
            var report = Report(Settings(), Sheet(("A1", "x", null)));

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = value });

            Assert.Equal("東京都千代田区\n丸の内1-1-1", result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void 空文字を書き込むとテンプレートの値を消せる()
        {
            var report = Report(Settings(), Sheet(("A1", "テンプレートの値", null)));

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = string.Empty });

            Assert.Equal(string.Empty, result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
            Assert.DoesNotContain(CellAddress.Parse("A1"), result.SubstitutedCells);
        }

        [Fact]
        public void 書き込んだセルを空文字で上書きすると書き込み済みの記録から外れる()
        {
            // 空文字は何も出力しないため、印刷範囲外・非表示の検出(要件2.13)の対象から外す(MarkSubstituted の削除の経路)。
            var report = Report(Settings(), Sheet(("A1", "x", null)));
            var written = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = "値" });
            Assert.Contains(CellAddress.Parse("A1"), written.SubstitutedCells);

            var cleared = _substitutor.ApplyCellOverrides(written, new Dictionary<string, string> { ["A1"] = string.Empty });

            Assert.DoesNotContain(CellAddress.Parse("A1"), cleared.SubstitutedCells);
        }

        [Fact]
        public void 空でない値を差し込んだセルだけが差し込み済みとして記録される()
        {
            var definition = Settings();
            var report = Report(definition, Sheet(("A1", "x", null), ("B1", "y", null), ("C1", "z", null)));

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = "値", ["B1"] = string.Empty });

            Assert.Contains(CellAddress.Parse("A1"), result.SubstitutedCells);
            Assert.DoesNotContain(CellAddress.Parse("B1"), result.SubstitutedCells);
            Assert.DoesNotContain(CellAddress.Parse("C1"), result.SubstitutedCells);
        }

        [Fact]
        public void セル番地直接指定でも制御文字はエラーになる()
        {
            var report = Report(Settings(), Sheet(("B5", "x", null)));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["B5"] = "a\tb" }));

            Assert.Equal("B5", ex.Target);
            Assert.Equal(CellAddress.Parse("B5"), ex.CellAddress);
        }

        [Fact]
        public void セル番地直接指定でも改行コードが統一され差し込み済みとして記録される()
        {
            var report = Report(Settings(), Sheet(("B5", "x", null)));

            var result = _substitutor.ApplyCellOverrides(
                report, new Dictionary<string, string> { ["B5"] = "1行目\r\n2行目" });

            Assert.Equal("1行目\n2行目", result.Sheet.GetCell(CellAddress.Parse("B5"))!.DisplayValue);
            Assert.Contains(CellAddress.Parse("B5"), result.SubstitutedCells);
        }

        [Theory]
        [InlineData("東京都\u2028丸の内")] // 行区切り
        [InlineData("東京都\u2029丸の内")] // 段落区切り
        public void Unicodeの行区切りと段落区切りもLFに統一される(string value)
        {
            var report = Report(Settings(), Sheet(("A1", "x", null)));

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = value });

            Assert.Equal("東京都\n丸の内", result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void 不正な値があるときは他のセルも書き換えない()
        {
            var definition = Settings();
            var report = Report(definition, Sheet(("A1", "もとのA", null), ("B1", "もとのB", null)));

            Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = "正常", ["B1"] = "不正\t" }));

            Assert.Equal("もとのA", report.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }
    }
}
