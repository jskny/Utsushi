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
        public void 改行以外の制御文字を含む値はエラーになる(string value, string expectedCodePoint)
        {
            var report = Report(Definition(Field("InvoiceNo", "C3")), Sheet(("C3", "x", null)));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.Apply(report, new Dictionary<string, string> { ["InvoiceNo"] = value }));

            Assert.Equal("InvoiceNo", ex.Target);
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
            var report = Report(Definition(Field("Name", "A1")), Sheet(("A1", "x", null)));

            Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.Apply(report, new Dictionary<string, string> { ["Name"] = value }));
        }

        [Fact]
        public void サロゲートペアの文字は受け付ける()
        {
            var report = Report(Definition(Field("Name", "A1")), Sheet(("A1", "x", null)));

            var result = _substitutor.Apply(report, new Dictionary<string, string> { ["Name"] = "𠮷野 様" });

            Assert.Equal("𠮷野 様", result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void nullの値はエラーになる()
        {
            var report = Report(Definition(Field("Name", "A1")), Sheet(("A1", "x", null)));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.Apply(report, new Dictionary<string, string> { ["Name"] = null! }));

            Assert.Equal("Name", ex.Target);
        }

        [Theory]
        [InlineData("東京都千代田区\r\n丸の内1-1-1")]
        [InlineData("東京都千代田区\r丸の内1-1-1")]
        [InlineData("東京都千代田区\n丸の内1-1-1")]
        public void 改行コードはLFに統一される(string value)
        {
            var report = Report(Definition(Field("Address", "A1")), Sheet(("A1", "x", null)));

            var result = _substitutor.Apply(report, new Dictionary<string, string> { ["Address"] = value });

            Assert.Equal("東京都千代田区\n丸の内1-1-1", result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("　")]
        public void 必須キーの値が空または空白のみならエラーになる(string value)
        {
            var report = Report(Definition(Field("CustomerName", "A1", required: true)), Sheet(("A1", "x", null)));

            var ex = Assert.Throws<RequiredSubstitutionValueMissingException>(
                () => _substitutor.Apply(report, new Dictionary<string, string> { ["CustomerName"] = value }));

            Assert.Equal("CustomerName", ex.Key);
            Assert.Contains("空", ex.Message);
        }

        [Fact]
        public void 任意キーに空文字を渡すとテンプレートの値を消せる()
        {
            var report = Report(Definition(Field("Remarks", "A1")), Sheet(("A1", "テンプレートの値", null)));

            var result = _substitutor.Apply(report, new Dictionary<string, string> { ["Remarks"] = string.Empty });

            Assert.Equal(string.Empty, result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
            Assert.DoesNotContain(CellAddress.Parse("A1"), result.SubstitutedCells);
        }

        [Fact]
        public void 空でない値を差し込んだセルだけが差し込み済みとして記録される()
        {
            var definition = Definition(Field("A", "A1"), Field("B", "B1"), Field("C", "C1"));
            var report = Report(definition, Sheet(("A1", "x", null), ("B1", "y", null), ("C1", "z", null)));

            var result = _substitutor.Apply(report, new Dictionary<string, string> { ["A"] = "値", ["B"] = string.Empty });

            Assert.Contains(CellAddress.Parse("A1"), result.SubstitutedCells);
            Assert.DoesNotContain(CellAddress.Parse("B1"), result.SubstitutedCells);
            Assert.DoesNotContain(CellAddress.Parse("C1"), result.SubstitutedCells);
        }

        [Fact]
        public void セル番地直接指定でも制御文字はエラーになる()
        {
            var report = Report(Definition(), Sheet(("B5", "x", null)));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["B5"] = "a\tb" }));

            Assert.Equal("B5", ex.Target);
            Assert.Equal(CellAddress.Parse("B5"), ex.CellAddress);
        }

        [Fact]
        public void セル番地直接指定でも改行コードが統一され差し込み済みとして記録される()
        {
            var report = Report(Definition(), Sheet(("B5", "x", null)));

            var result = _substitutor.ApplyCellOverrides(
                report, new Dictionary<string, string> { ["B5"] = "1行目\r\n2行目" });

            Assert.Equal("1行目\n2行目", result.Sheet.GetCell(CellAddress.Parse("B5"))!.DisplayValue);
            Assert.Contains(CellAddress.Parse("B5"), result.SubstitutedCells);
        }

        [Fact]
        public void セル番地直接指定で空文字にすると差し込み済みの記録から外れる()
        {
            var report = Report(Definition(Field("Name", "A1")), Sheet(("A1", "x", null)));
            var applied = _substitutor.Apply(report, new Dictionary<string, string> { ["Name"] = "値" });

            var result = _substitutor.ApplyCellOverrides(applied, new Dictionary<string, string> { ["A1"] = string.Empty });

            Assert.DoesNotContain(CellAddress.Parse("A1"), result.SubstitutedCells);
        }

        [Fact]
        public void 不正な値があるときは他のセルも書き換えない()
        {
            var definition = Definition(Field("A", "A1"), Field("B", "B1"));
            var report = Report(definition, Sheet(("A1", "もとのA", null), ("B1", "もとのB", null)));

            Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.Apply(report, new Dictionary<string, string> { ["A"] = "正常", ["B"] = "不正\t" }));

            Assert.Equal("もとのA", report.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }
    }
}
