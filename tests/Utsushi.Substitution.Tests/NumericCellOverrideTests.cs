using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Xunit;
using static Utsushi.Substitution.Tests.TestReports;

namespace Utsushi.Substitution.Tests
{
    /// <summary>数値の直接指定(要件14.3)。</summary>
    public sealed class NumericCellOverrideTests
    {
        private static readonly CellStyle Currency = CellStyle.Default with { NumberFormat = "\"¥\"#,##0;[Red]\"¥\"-#,##0" };

        private readonly CellSubstitutor _substitutor = new();

        [Fact]
        public void セルを数値のセルにしてセルの数値書式で表示する()
        {
            var report = Report(Definition(), Sheet(("B2", "テンプレートの値", Currency)));

            var result = _substitutor.ApplyNumericCellOverrides(report, new Dictionary<string, double> { ["b2"] = -1234.4 });

            var cell = result.Sheet.Cells[CellAddress.Parse("B2")];
            Assert.Equal(CellValueKind.Number, cell.ValueKind);
            Assert.Equal("¥-1,234", cell.DisplayValue);
            Assert.Equal(new ArgbColor(0xFF, 0xFF, 0x00, 0x00), cell.FormatColor);
            Assert.Same(Currency, cell.Style);
            Assert.Contains(CellAddress.Parse("B2"), result.SubstitutedCells);
            Assert.Contains(CellAddress.Parse("B2"), result.OverriddenCells);
        }

        [Fact]
        public void セルが無い位置は位置の書式の数値書式で表示する()
        {
            var sheet = Sheet(("A1", "見出し", null)) with
            {
                ColumnStyles = new[] { new ColumnStyleRange(3, 3, Currency) },
            };

            var result = _substitutor.ApplyNumericCellOverrides(
                Report(Definition(), sheet), new Dictionary<string, double> { ["C5"] = 1000 });

            Assert.Equal("¥1,000", result.Sheet.Cells[CellAddress.Parse("C5")].DisplayValue);
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        public void 表示できない数値はエラーにする(double value)
        {
            var report = Report(Definition(), Sheet(("A1", "x", null)));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyNumericCellOverrides(report, new Dictionary<string, double> { ["A1"] = value }));
            Assert.Equal("A1", ex.Target);
        }

        [Fact]
        public void 件数の上限は文字列の直接指定と合わせて数える()
        {
            var overridden = new HashSet<CellAddress>();
            for (var i = 1; i <= CellSubstitutor.MaxCellOverrideCount; i++)
            {
                overridden.Add(new CellAddress(i, 1));
            }

            var report = Report(Definition(), Sheet(("A1", "x", null))) with { OverriddenCells = overridden };

            var ex = Assert.Throws<InvalidSubstitutionValueException>(
                () => _substitutor.ApplyNumericCellOverrides(report, new Dictionary<string, double> { ["B1"] = 1 }));
            Assert.Equal("cellOverrides", ex.Target);
        }

        [Fact]
        public void 番地の検証は文字列の直接指定と同じ()
        {
            var report = Report(Definition(), Sheet(("A1", "x", null)));

            Assert.Throws<InvalidCellOverrideAddressException>(
                () => _substitutor.ApplyNumericCellOverrides(report, new Dictionary<string, double> { ["1A"] = 1 }));
        }
    }
}
