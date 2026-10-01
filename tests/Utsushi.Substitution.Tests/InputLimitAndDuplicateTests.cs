using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Xunit;
using static Utsushi.Substitution.Tests.TestReports;

namespace Utsushi.Substitution.Tests
{
    /// <summary>
    /// 差し込み値・セル上書き値の長さと件数の上限、同じセルへの重複した直接指定の検出、
    /// ASCII以外の英字を含むセル番地の拒否を確かめる。
    /// </summary>
    public sealed class InputLimitAndDuplicateTests
    {
        private readonly CellSubstitutor _substitutor = new();

        [Fact]
        public void 上限ちょうどの長さの差し込み値は受け付ける()
        {
            var report = Report(Definition(Field("Name", "A1")), Sheet(("A1", "x", null)));
            var value = new string('あ', CellSubstitutor.MaxValueLength);

            var result = _substitutor.Apply(report, new Dictionary<string, string> { ["Name"] = value });

            Assert.Equal(value, result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void 上限を超える長さの差し込み値はエラーになる()
        {
            var report = Report(Definition(Field("Name", "A1")), Sheet(("A1", "x", null)));
            var value = new string('あ', CellSubstitutor.MaxValueLength + 1);

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.Apply(report, new Dictionary<string, string> { ["Name"] = value }));

            Assert.Equal("Name", ex.Target);
            Assert.Equal(CellAddress.Parse("A1"), ex.CellAddress);
            Assert.Equal(ProcessingStage.Substitution, ex.Stage);
            Assert.Contains(CellSubstitutor.MaxValueLength.ToString(System.Globalization.CultureInfo.InvariantCulture), ex.Message);
        }

        [Fact]
        public void 上限を超える長さのセル上書き値はエラーになる()
        {
            var report = Report(Definition(), Sheet(("B2", "x", null)));
            var value = new string('a', CellSubstitutor.MaxValueLength + 1);

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["B2"] = value }));

            Assert.Equal("B2", ex.Target);
            Assert.Equal(CellAddress.Parse("B2"), ex.CellAddress);
        }

        [Fact]
        public void 上限ちょうどの件数のセル上書きは受け付ける()
        {
            var report = Report(Definition(), Sheet());
            var overrides = Overrides(CellSubstitutor.MaxCellOverrideCount);

            var result = _substitutor.ApplyCellOverrides(report, overrides);

            Assert.Equal(CellSubstitutor.MaxCellOverrideCount, result.SubstitutedCells.Count);
        }

        [Fact]
        public void 上限を超える件数のセル上書きはエラーになる()
        {
            var report = Report(Definition(), Sheet());
            var overrides = Overrides(CellSubstitutor.MaxCellOverrideCount + 1);

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyCellOverrides(report, overrides));

            Assert.Equal(ProcessingStage.Substitution, ex.Stage);
            Assert.Equal("test-report", ex.ReportCode);
        }

        [Theory]
        [InlineData("B1", "b1")]
        [InlineData("B1", "$B$1")]
        [InlineData("b1", "$b1")]
        [InlineData("B1", " B1 ")]
        public void 表記が違っても同じセルを指す直接指定はエラーになる(string first, string second)
        {
            var report = Report(Definition(), Sheet(("B1", "x", null)));
            var overrides = new Dictionary<string, string> { [first] = "1つ目", [second] = "2つ目" };

            var ex = Assert.Throws<InvalidCellOverrideAddressException>(
                () => _substitutor.ApplyCellOverrides(report, overrides));

            Assert.Contains("B1", ex.Message);
            Assert.Equal(ProcessingStage.Substitution, ex.Stage);
        }

        [Theory]
        [InlineData("é1")] // char.IsLetter は真だが、列番号に換算すると EG 列になっていた
        [InlineData("Ａ1")] // 全角英字
        [InlineData("Aé1")]
        [InlineData("ß1")]
        public void ASCII以外の英字を含むセル番地はエラーになる(string address)
        {
            Assert.False(CellAddress.TryParse(address, out _));
            Assert.False(CellAddress.TryParseColumnName(address.Substring(0, address.Length - 1), out _));

            var report = Report(Definition(), Sheet(("A1", "x", null)));
            var ex = Assert.Throws<InvalidCellOverrideAddressException>(
                () => _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { [address] = "v" }));

            Assert.Equal(address, ex.Address);
        }

        [Theory]
        [InlineData("a1", 1, 1)]
        [InlineData("$XFD$1048576", 1048576, 16384)]
        [InlineData("zz10", 10, 702)]
        public void ASCII英字のセル番地は大文字小文字を問わず解釈できる(string text, int row, int column)
        {
            Assert.True(CellAddress.TryParse(text, out var address));
            Assert.Equal(new CellAddress(row, column), address);
        }

        private static Dictionary<string, string> Overrides(int count)
        {
            var overrides = new Dictionary<string, string>(count);
            for (var i = 0; i < count; i++)
            {
                // 1列目から順に、1行あたり100セルずつ埋める。
                var address = new CellAddress((i / 100) + 1, (i % 100) + 1);
                overrides[address.ToString()] = "v";
            }

            return overrides;
        }
    }
}
