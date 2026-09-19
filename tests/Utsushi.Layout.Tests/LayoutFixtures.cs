using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Layout.Tests;

/// <summary>レイアウト検証用のモデルを組み立てるヘルパー。</summary>
internal static class LayoutFixtures
{
    /// <summary>全セルが同じ幅・高さの単純なシートを作る。</summary>
    public static SheetModel UniformSheet(
        int rows,
        int columns,
        double columnWidth = 10.0,
        double rowHeightPt = 20.0,
        PageSetupModel? pageSetup = null,
        IEnumerable<CellRange>? mergedRanges = null,
        CellStyle? style = null)
    {
        var cells = new Dictionary<CellAddress, CellModel>();
        for (var r = 1; r <= rows; r++)
        {
            for (var c = 1; c <= columns; c++)
            {
                var address = new CellAddress(r, c);
                cells[address] = new CellModel(
                    address.ToString(), CellValueKind.Text, style ?? CellStyle.Default, address.ToString());
            }
        }

        return new SheetModel(
            "テストシート",
            cells,
            (mergedRanges ?? Enumerable.Empty<CellRange>()).Select(r => new MergedRange(r)).ToList(),
            Enumerable.Repeat(columnWidth, columns).ToList(),
            Enumerable.Repeat(rowHeightPt, rows).ToList(),
            columnWidth,
            rowHeightPt,
            new HashSet<int>(),
            new HashSet<int>(),
            pageSetup ?? PageSetupModel.Default);
    }

    public static ReportDefinition Definition(
        double maxDigitWidthPx = ReportDefinition.DefaultMaxDigitWidthPx,
        CellRange? printAreaOverride = null,
        IEnumerable<SubstitutionFieldDefinition>? fields = null) =>
        new(
            "test-report",
            "テストシート",
            (fields ?? Enumerable.Empty<SubstitutionFieldDefinition>()).ToList(),
            ReportDefinition.DefaultToleranceMm,
            UnsupportedElementPolicy.Ignore,
            maxDigitWidthPx,
            printAreaOverride);

    /// <summary>A4縦・余白ゼロのページ設定(印字可能領域の計算を単純化するため)。</summary>
    public static PageSetupModel NoMarginA4(
        IReadOnlyList<CellRange>? printAreas = null,
        IReadOnlyList<int>? rowBreaks = null,
        IReadOnlyList<int>? columnBreaks = null,
        PrintTitles? printTitles = null,
        PageOrder pageOrder = PageOrder.DownThenOver,
        PageScaling? scaling = null) =>
        PageSetupModel.Default with
        {
            Margins = new PageMargins(0, 0, 0, 0, 0, 0),
            PrintAreas = printAreas ?? new List<CellRange>(),
            ManualRowBreaks = rowBreaks ?? new List<int>(),
            ManualColumnBreaks = columnBreaks ?? new List<int>(),
            PrintTitles = printTitles ?? PrintTitles.None,
            PageOrder = pageOrder,
            Scaling = scaling ?? PageScaling.Normal,
        };

    public static IEnumerable<TextCommand> Texts(PageLayout page) => page.Commands.OfType<TextCommand>();

    public static IEnumerable<LineCommand> Lines(PageLayout page) => page.Commands.OfType<LineCommand>();

    public static IEnumerable<FillRectCommand> Fills(PageLayout page) => page.Commands.OfType<FillRectCommand>();
}
