using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using static Utsushi.Substitution.Tests.TestReports;

namespace Utsushi.Substitution.Tests;

/// <summary>
/// セル値置換の検証(要件2.1-2.5、タスク4.1-4.3 / 8.1)。
/// </summary>
public sealed class CellSubstitutorTests
{
    private readonly CellSubstitutor _substitutor = new();

    [Fact]
    public void 帳票定義に登録されたセルの値だけを置換する()
    {
        var definition = Definition(Field("InvoiceNo", "C3"));
        var sheet = Sheet(("C3", "INV-0000", null), ("C4", "触らないセル", null));
        var report = Report(definition, sheet);

        var result = _substitutor.Apply(report, new Dictionary<string, string> { ["InvoiceNo"] = "INV-2026-1" });

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

        var definition = Definition(Field("Amount", "D5"));
        var report = Report(definition, Sheet(("D5", "0", style)));

        var result = _substitutor.Apply(report, new Dictionary<string, string> { ["Amount"] = "1,234" });

        var cell = result.Sheet.GetCell(CellAddress.Parse("D5"))!;
        Assert.Equal("1,234", cell.DisplayValue);
        Assert.Same(style, cell.Style);
    }

    [Fact]
    public void 入力のモデルは変更されない()
    {
        var definition = Definition(Field("Key", "A1"));
        var report = Report(definition, Sheet(("A1", "もとの値", null)));

        _substitutor.Apply(report, new Dictionary<string, string> { ["Key"] = "あたらしい値" });

        Assert.Equal("もとの値", report.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
    }

    [Fact]
    public void 未知の置換キーはエラーになる()
    {
        var definition = Definition(Field("Known", "A1"));
        var report = Report(definition, Sheet(("A1", "x", null)));

        var ex = Assert.Throws<SubstitutionKeyNotFoundException>(
            () => _substitutor.Apply(report, new Dictionary<string, string> { ["Unknown"] = "v" }));

        Assert.Equal("Unknown", ex.Key);
        Assert.Equal("test-report", ex.ReportCode);
        Assert.Equal(ProcessingStage.Substitution, ex.Stage);
    }

    [Fact]
    public void 必須キーの値が渡されなければエラーになる()
    {
        var definition = Definition(
            Field("Required", "A1", required: true),
            Field("Optional", "A2"));
        var report = Report(definition, Sheet(("A1", "x", null), ("A2", "y", null)));

        var ex = Assert.Throws<RequiredSubstitutionValueMissingException>(
            () => _substitutor.Apply(report, new Dictionary<string, string> { ["Optional"] = "v" }));

        Assert.Equal("Required", ex.Key);
        Assert.Equal(CellAddress.Parse("A1"), ex.CellAddress);
    }

    [Fact]
    public void 任意キーの値が渡されなければテンプレートの値が残る()
    {
        var definition = Definition(Field("Optional", "A1"));
        var report = Report(definition, Sheet(("A1", "テンプレートの値", null)));

        var result = _substitutor.Apply(report, new Dictionary<string, string>());

        Assert.Equal("テンプレートの値", result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
    }

    [Fact]
    public void 置換したセルのはみ出し挙動がLayoutへ引き渡される()
    {
        var definition = Definition(
            Field("Shrink", "A1", overflow: OverflowBehavior.Shrink),
            Field("Wrap", "A2", overflow: OverflowBehavior.Wrap));
        var report = Report(definition, Sheet(("A1", "a", null), ("A2", "b", null)));

        var result = _substitutor.Apply(
            report, new Dictionary<string, string> { ["Shrink"] = "x", ["Wrap"] = "y" });

        Assert.Equal(OverflowBehavior.Shrink, result.GetOverflowBehavior(CellAddress.Parse("A1")));
        Assert.Equal(OverflowBehavior.Wrap, result.GetOverflowBehavior(CellAddress.Parse("A2")));
    }

    [Fact]
    public void 置換していないセルにははみ出し挙動を設定しない()
    {
        var definition = Definition(Field("Key", "A1", overflow: OverflowBehavior.Clip));
        var report = Report(definition, Sheet(("A1", "a", null), ("B2", "b", null)));

        var result = _substitutor.Apply(report, new Dictionary<string, string>());

        Assert.Null(result.GetOverflowBehavior(CellAddress.Parse("A1")));
        Assert.Null(result.GetOverflowBehavior(CellAddress.Parse("B2")));
    }

    [Fact]
    public void 空文字への置換もできる()
    {
        var definition = Definition(Field("Key", "A1"));
        var report = Report(definition, Sheet(("A1", "消したい値", null)));

        var result = _substitutor.Apply(report, new Dictionary<string, string> { ["Key"] = string.Empty });

        Assert.Equal(string.Empty, result.Sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
    }

    [Fact]
    public void 数式セルを置換すると数式ではなく文字列になる()
    {
        var definition = Definition(Field("Key", "A1"));
        var sheet = Sheet();
        var cells = new Dictionary<CellAddress, CellModel>
        {
            [CellAddress.Parse("A1")] = new("100", CellValueKind.Number, CellStyle.Default, "100", HasFormula: true),
        };
        var report = Report(definition, sheet with { Cells = cells });

        var result = _substitutor.Apply(report, new Dictionary<string, string> { ["Key"] = "差し込み値" });

        var cell = result.Sheet.GetCell(CellAddress.Parse("A1"))!;
        Assert.False(cell.HasFormula);
        Assert.Equal(CellValueKind.Text, cell.ValueKind);
        Assert.Equal("差し込み値", cell.DisplayValue);
    }
}
