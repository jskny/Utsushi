using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Xunit;
using static Utsushi.Substitution.Tests.TestReports;

namespace Utsushi.Substitution.Tests
{
    /// <summary>
    /// ファイルにセルが無い位置への置換は、行・列・ブックの標準の書式でセルを作る(要件1.10)。
    /// </summary>
    public sealed class EffectiveStyleSubstitutionTests
    {
        private static readonly CellStyle Standard =
            CellStyle.Default with { Font = FontStyle.Default with { Name = "游ゴシック" } };

        private static readonly CellStyle RowStyle =
            CellStyle.Default with { Font = FontStyle.Default with { Name = "ＭＳ 明朝" } };

        private static readonly CellStyle ColumnStyle =
            CellStyle.Default with { Font = FontStyle.Default with { Name = "メイリオ", Bold = true } };

        private static SheetModel SheetWithPositionStyles() =>
            Sheet(("A1", "見出し", null)) with
            {
                DefaultCellStyle = Standard,
                RowStyles = new Dictionary<int, CellStyle> { [3] = RowStyle },
                ColumnStyles = new[] { new ColumnStyleRange(2, 2, ColumnStyle) },
            };

        [Theory]
        [InlineData("C2", "游ゴシック")]   // 行・列の書式なし → ブックの標準の書式
        [InlineData("B2", "メイリオ")]     // 列の書式
        [InlineData("B3", "ＭＳ 明朝")]    // 行の書式は列の書式より優先
        public void 置換キーでセルが無い位置に差し込むと位置の書式を使う(string cell, string expectedFont)
        {
            var report = Report(Definition(Field("Key", cell)), SheetWithPositionStyles());

            var result = new CellSubstitutor().Apply(report, new Dictionary<string, string> { ["Key"] = "値" });

            Assert.Equal(expectedFont, result.Sheet.Cells[CellAddress.Parse(cell)].Style.Font.Name);
        }

        [Fact]
        public void セル番地直接指定でセルが無い位置に差し込むと位置の書式を使う()
        {
            var report = Report(Definition(), SheetWithPositionStyles());

            var result = new CellSubstitutor().ApplyCellOverrides(report, new Dictionary<string, string> { ["B5"] = "値" });

            Assert.Equal("メイリオ", result.Sheet.Cells[CellAddress.Parse("B5")].Style.Font.Name);
        }
    }
}
