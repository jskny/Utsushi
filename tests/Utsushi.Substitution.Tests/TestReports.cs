using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Substitution.Tests
{
    /// <summary>テスト用の帳票モデルを組み立てるヘルパー。</summary>
    internal static class TestReports
    {
        public static ReportDefinition Definition() =>
            new(DocumentName: "test-report", SheetName: "テストシート", MaxDigitWidthPx: ReportDefinition.DefaultMaxDigitWidthPx);

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
                PageSetup: PageSetupModel.Default,
                DrawingObjects: Array.Empty<DrawingObjectModel>());
        }

        public static ReportModel Report(ReportDefinition definition, SheetModel sheet) =>
            ReportModel.Create(definition, sheet);
    }
}
