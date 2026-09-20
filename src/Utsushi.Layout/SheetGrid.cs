using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout;

/// <summary>
/// シートの列幅・行高をポイントへ換算し、印刷対象の行/列を並べた格子。
/// </summary>
/// <remarks>非表示行・非表示列は印刷されないため、ここで除外する。</remarks>
public sealed class SheetGrid
{
    private readonly Dictionary<int, double> _columnWidthPt;
    private readonly Dictionary<int, double> _rowHeightPt;

    private SheetGrid(
        IReadOnlyList<int> columns,
        IReadOnlyList<int> rows,
        Dictionary<int, double> columnWidthPt,
        Dictionary<int, double> rowHeightPt,
        CellRange printRange)
    {
        Columns = columns;
        Rows = rows;
        _columnWidthPt = columnWidthPt;
        _rowHeightPt = rowHeightPt;
        PrintRange = printRange;
    }

    /// <summary>印刷対象の列番号(昇順、非表示列を除く)。</summary>
    public IReadOnlyList<int> Columns { get; }

    /// <summary>印刷対象の行番号(昇順、非表示行を除く)。</summary>
    public IReadOnlyList<int> Rows { get; }

    /// <summary>印刷対象のセル範囲(非表示行/列を含む矩形)。</summary>
    public CellRange PrintRange { get; }

    /// <summary>
    /// 印刷範囲でクリッピングした格子を作る(要件3.1)。
    /// </summary>
    /// <param name="sheet">対象シート。</param>
    /// <param name="printRange">印刷範囲。</param>
    /// <param name="maxDigitWidthPx">列幅換算に用いる標準フォントの最大数字幅(ピクセル)。</param>
    /// <param name="printTitles">
    /// 印刷タイトル(要件3.4)。Excelでは印刷タイトルの行/列は印刷範囲(Print_Area)と独立に指定でき、
    /// 印刷範囲に含まれない行/列でも各ページに繰り返し出力される。そのため印刷範囲だけでなく
    /// 印刷タイトルが指す行/列も併せて格子に含める(印刷範囲とタイトルの間にある行/列は含めない)。
    /// </param>
    public static SheetGrid Create(
        SheetModel sheet, CellRange printRange, double maxDigitWidthPx, PrintTitles? printTitles = null)
    {
        var columnRanges = new List<(int First, int Last)> { (printRange.FirstColumn, printRange.LastColumn) };
        if (printTitles is { HasColumns: true })
        {
            columnRanges.Add((printTitles.FirstColumn!.Value, printTitles.LastColumn!.Value));
        }

        var columns = new List<int>();
        var columnWidths = new Dictionary<int, double>();
        foreach (var c in CollectDistinctIndices(columnRanges))
        {
            if (sheet.IsColumnHidden(c))
            {
                continue;
            }

            var widthPt = ExcelUnitConverter.ColumnWidthToPoints(sheet.GetColumnWidth(c), maxDigitWidthPx);
            if (widthPt <= 0)
            {
                // 幅0の列は非表示と同じ扱い。
                continue;
            }

            columns.Add(c);
            columnWidths[c] = widthPt;
        }

        var rowRanges = new List<(int First, int Last)> { (printRange.FirstRow, printRange.LastRow) };
        if (printTitles is { HasRows: true })
        {
            rowRanges.Add((printTitles.FirstRow!.Value, printTitles.LastRow!.Value));
        }

        var rows = new List<int>();
        var rowHeights = new Dictionary<int, double>();
        foreach (var r in CollectDistinctIndices(rowRanges))
        {
            if (sheet.IsRowHidden(r))
            {
                continue;
            }

            var heightPt = sheet.GetRowHeight(r);
            if (heightPt <= 0)
            {
                continue;
            }

            rows.Add(r);
            rowHeights[r] = heightPt;
        }

        return new SheetGrid(columns, rows, columnWidths, rowHeights, printRange);
    }

    /// <summary>複数区間の索引を、重複と順序を整えたうえで昇順に列挙する。</summary>
    private static IEnumerable<int> CollectDistinctIndices(IReadOnlyList<(int First, int Last)> ranges)
    {
        var indices = new SortedSet<int>();
        foreach (var (first, last) in ranges)
        {
            for (var i = first; i <= last; i++)
            {
                indices.Add(i);
            }
        }

        return indices;
    }

    /// <summary>列幅(ポイント)。印刷対象外の列は0。</summary>
    public double GetColumnWidthPt(int column) =>
        _columnWidthPt.TryGetValue(column, out var width) ? width : 0.0;

    /// <summary>行高(ポイント)。印刷対象外の行は0。</summary>
    public double GetRowHeightPt(int row) =>
        _rowHeightPt.TryGetValue(row, out var height) ? height : 0.0;

    /// <summary>指定した行の並びの合計高さ(ポイント)。</summary>
    public double SumRowHeights(IEnumerable<int> rows)
    {
        var total = 0.0;
        foreach (var row in rows)
        {
            total += GetRowHeightPt(row);
        }

        return total;
    }

    /// <summary>指定した列の並びの合計幅(ポイント)。</summary>
    public double SumColumnWidths(IEnumerable<int> columns)
    {
        var total = 0.0;
        foreach (var column in columns)
        {
            total += GetColumnWidthPt(column);
        }

        return total;
    }
}
