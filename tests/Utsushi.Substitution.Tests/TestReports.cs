using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Substitution.Tests;

/// <summary>テスト用の帳票モデルを組み立てるヘルパー。</summary>
internal static class TestReports
{
    public static ReportDefinition Definition(params SubstitutionFieldDefinition[] fields) =>
        new(
            ReportCode: "test-report",
            SheetName: "テストシート",
            SubstitutionFields: fields,
            ToleranceMm: ReportDefinition.DefaultToleranceMm,
            UnsupportedElements: UnsupportedElementPolicy.Ignore,
            MaxDigitWidthPx: ReportDefinition.DefaultMaxDigitWidthPx,
            PrintAreaOverride: null);

    public static SubstitutionFieldDefinition Field(
        string key, string cell, bool required = false, OverflowBehavior? overflow = null) =>
        new(key, CellAddress.Parse(cell), required, overflow);

    /// <summary>指定セルに値と書式を持つシートを作る。</summary>
    public static SheetModel Sheet(params (string Cell, string? Value, CellStyle? Style)[] cells)
    {
        var map = new Dictionary<CellAddress, CellModel>();
        foreach (var (cell, value, style) in cells)
        {
            var address = CellAddress.Parse(cell);
            map[address] = new CellModel(
                value,
                value is null ? CellValueKind.Blank : CellValueKind.Text,
                style ?? CellStyle.Default,
                value);
        }

        return new SheetModel(
            "テストシート",
            map,
            new List<MergedRange>(),
            new List<double>(),
            new List<double>(),
            DefaultColumnWidth: 8.43,
            DefaultRowHeight: 15.0,
            HiddenColumns: new HashSet<int>(),
            HiddenRows: new HashSet<int>(),
            PageSetup: PageSetupModel.Default);
    }

    public static ReportModel Report(ReportDefinition definition, SheetModel sheet) =>
        ReportModel.Create(definition, sheet);
}
