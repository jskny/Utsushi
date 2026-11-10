using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using static Utsushi.Substitution.Tests.TestReports;

namespace Utsushi.Substitution.Tests
{
    /// <summary>
    /// セル値置換の検証(要件2.1-2.5、タスク4.1-4.3 / 8.1)。
    /// </summary>
    public sealed class CellSubstitutorTests
    {
        private readonly CellSubstitutor _substitutor = new();

        [Fact]
        public void 指定したセルの値だけを書き換える()
        {
            var definition = Settings();
            var sheet = Sheet(("C3", "INV-0000", null), ("C4", "触らないセル", null));
            var report = Report(definition, sheet);

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["C3"] = "INV-2026-1" });

            Assert.Equal("INV-2026-1", result.Sheet.GetCell(CellAddress.Parse("C3"))!.DisplayValue);
            Assert.Equal("触らないセル", result.Sheet.GetCell(CellAddress.Parse("C4"))!.DisplayValue);
        }

        [Fact]
        public void 置換してもセルの書式は変わらない()
        {
            var style = CellStyle.Default with
            {
                Font = FontStyle.Default with { Bold = true, SizePt = 14.0 },
                HAlign = HorizontalAlignment.Right,
                NumberFormat = "#,##0",
                Borders = new BorderSet(
                    new BorderEdge(BorderLineStyle.Thin, ArgbColor.Black),
                    BorderEdge.None, BorderEdge.None, BorderEdge.None, BorderEdge.None, BorderEdge.None),
            };

            var definition = Settings();
            var report = Report(definition, Sheet(("D5", "0", style)));

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["D5"] = "1,234" });

            var cell = result.Sheet.GetCell(CellAddress.Parse("D5"))!;
            Assert.Equal("1,234", cell.DisplayValue);
            Assert.Same(style, cell.Style);
        }

        [Fact]
        public void 入力のモデルは変更されない()
        {
            var definition = Settings();
            var report = Report(definition, Sheet(("A1", "もとの値", null)));

            _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = "あたらしい値" });

            Assert.Equal("もとの値", report.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void 空文字への置換もできる()
        {
            var definition = Settings();
            var report = Report(definition, Sheet(("A1", "消したい値", null)));

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = string.Empty });

            Assert.Equal(string.Empty, result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void 数式セルを置換すると数式ではなく文字列になる()
        {
            var definition = Settings();
            var sheet = Sheet();
            var cells = new Dictionary<CellAddress, CellModel>
            {
                [CellAddress.Parse("A1")] = new("100", CellValueKind.Number, CellStyle.Default, "100", HasFormula: true),
            };
            var report = Report(definition, sheet with { Cells = cells });

            var result = _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = "差し込み値" });

            var cell = result.Sheet.GetCell(CellAddress.Parse("A1"))!;
            Assert.False(cell.HasFormula);
            Assert.Equal(CellValueKind.Text, cell.ValueKind);
            Assert.Equal("差し込み値", cell.DisplayValue);
        }

        // --- ApplyCellOverrides(セル番地直接指定、要件2.7, 2.8) ---

        [Fact]
        public void テンプレートの値の無いセルにもセル番地指定で書き込める()
        {
            var definition = Settings();
            var sheet = Sheet(("C3", "INV-0000", null), ("B5", "もとの備考", null));
            var report = Report(definition, sheet);

            var result = _substitutor.ApplyCellOverrides(
                report, new Dictionary<string, string> { ["B5"] = "臨時の備考" });

            Assert.Equal("臨時の備考", result.Sheet.GetCell(CellAddress.Parse("B5"))!.DisplayValue);
            Assert.Equal("INV-0000", result.Sheet.GetCell(CellAddress.Parse("C3"))!.DisplayValue);
        }

        [Fact]
        public void セル番地指定の上書きでも書式は変わらない()
        {
            var style = CellStyle.Default with { Font = FontStyle.Default with { Bold = true } };
            var definition = Settings();
            var report = Report(definition, Sheet(("D5", "0", style)));

            var result = _substitutor.ApplyCellOverrides(
                report, new Dictionary<string, string> { ["D5"] = "1,234" });

            var cell = result.Sheet.GetCell(CellAddress.Parse("D5"))!;
            Assert.Equal("1,234", cell.DisplayValue);
            Assert.Same(style, cell.Style);
        }

        [Fact]
        public void セル番地指定の上書きでも入力のモデルは変更されない()
        {
            var definition = Settings();
            var report = Report(definition, Sheet(("A1", "もとの値", null)));

            _substitutor.ApplyCellOverrides(report, new Dictionary<string, string> { ["A1"] = "あたらしい値" });

            Assert.Equal("もとの値", report.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
        }

        [Fact]
        public void 存在しなかったセルへのセル番地指定も新規セルとして追加される()
        {
            var definition = Settings();
            var report = Report(definition, Sheet());

            var result = _substitutor.ApplyCellOverrides(
                report, new Dictionary<string, string> { ["Z9"] = "新規の値" });

            Assert.Equal("新規の値", result.Sheet.GetCell(CellAddress.Parse("Z9"))!.DisplayValue);
        }

        [Fact]
        public void A1形式として解釈できないセル番地はエラーになる()
        {
            var definition = Settings();
            var report = Report(definition, Sheet());

            var ex = Assert.Throws<InvalidCellOverrideAddressException>(
                () => _substitutor.ApplyCellOverrides(
                    report, new Dictionary<string, string> { ["not-a-cell"] = "v" }));

            Assert.Equal("not-a-cell", ex.Address);
            Assert.Equal("test-report", ex.DocumentName);
            Assert.Equal(ProcessingStage.Substitution, ex.Stage);
        }

        [Fact]
        public void 結合セルのアンカーへの上書きはできる()
        {
            var definition = Settings();
            var sheet = Sheet(("B5", "もとの値", null)) with
            {
                MergedRanges = new List<MergedRange> { new(CellRange.Parse("B5:D5")) },
            };
            var report = Report(definition, sheet);

            var result = _substitutor.ApplyCellOverrides(
                report, new Dictionary<string, string> { ["B5"] = "上書き後" });

            Assert.Equal("上書き後", result.Sheet.GetCell(CellAddress.Parse("B5"))!.DisplayValue);
        }

        [Fact]
        public void 結合セルの非アンカーへの直接指定はエラーになる()
        {
            var definition = Settings();
            var sheet = Sheet(("B5", "もとの値", null)) with
            {
                MergedRanges = new List<MergedRange> { new(CellRange.Parse("B5:D5")) },
            };
            var report = Report(definition, sheet);

            var ex = Assert.Throws<NonAnchorMergedCellOverrideException>(
                () => _substitutor.ApplyCellOverrides(
                    report, new Dictionary<string, string> { ["D5"] = "見えなくなる値" }));

            Assert.Equal(CellAddress.Parse("D5"), ex.CellAddress);
            Assert.Equal(CellAddress.Parse("B5"), ex.AnchorAddress);
            Assert.Equal("test-report", ex.DocumentName);
        }
    }
}
