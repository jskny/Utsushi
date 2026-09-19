using System;
using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Parsing.Model;

/// <summary>
/// 1シート分の内部モデル。OOXML の型を一切含まない POCO。
/// </summary>
/// <param name="Name">シート名。</param>
/// <param name="Cells">値または書式を持つセル。空セルは含まれない場合がある。</param>
/// <param name="MergedRanges">結合セル範囲。</param>
/// <param name="ColumnWidths">列幅。索引0が列A。単位はExcelの「文字数」単位。</param>
/// <param name="RowHeights">行高。索引0が行1。単位はポイント。</param>
/// <param name="DefaultColumnWidth"><paramref name="ColumnWidths"/> の範囲外の列に適用する既定列幅(文字数単位)。</param>
/// <param name="DefaultRowHeight"><paramref name="RowHeights"/> の範囲外の行に適用する既定行高(ポイント)。</param>
/// <param name="HiddenColumns">非表示の列番号(1始まり)。</param>
/// <param name="HiddenRows">非表示の行番号(1始まり)。</param>
/// <param name="PageSetup">ページ設定。</param>
public sealed record SheetModel(
    string Name,
    IReadOnlyDictionary<CellAddress, CellModel> Cells,
    IReadOnlyList<MergedRange> MergedRanges,
    IReadOnlyList<double> ColumnWidths,
    IReadOnlyList<double> RowHeights,
    double DefaultColumnWidth,
    double DefaultRowHeight,
    IReadOnlySet<int> HiddenColumns,
    IReadOnlySet<int> HiddenRows,
    PageSetupModel PageSetup)
{
    /// <summary>指定セルを取得する。存在しない場合は null。</summary>
    public CellModel? GetCell(CellAddress address) =>
        Cells.TryGetValue(address, out var cell) ? cell : null;

    /// <summary>1始まりの列番号に対する列幅(文字数単位)。定義が無い列には既定値を返す。</summary>
    public double GetColumnWidth(int column)
    {
        if (column < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(column), column, "列番号は1始まりです。");
        }

        return column <= ColumnWidths.Count ? ColumnWidths[column - 1] : DefaultColumnWidth;
    }

    /// <summary>1始まりの行番号に対する行高(ポイント)。定義が無い行には既定値を返す。</summary>
    public double GetRowHeight(int row)
    {
        if (row < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(row), row, "行番号は1始まりです。");
        }

        return row <= RowHeights.Count ? RowHeights[row - 1] : DefaultRowHeight;
    }

    public bool IsColumnHidden(int column) => HiddenColumns.Contains(column);

    public bool IsRowHidden(int row) => HiddenRows.Contains(row);

    /// <summary>指定セルを含む結合範囲を返す。結合されていない場合は null。</summary>
    public MergedRange? FindMergedRange(CellAddress address)
    {
        for (var i = 0; i < MergedRanges.Count; i++)
        {
            if (MergedRanges[i].Contains(address))
            {
                return MergedRanges[i];
            }
        }

        return null;
    }

    /// <summary>セルが存在する範囲(使用範囲)。セルが1つも無い場合は null。</summary>
    public CellRange? GetUsedRange()
    {
        var hasAny = false;
        int minRow = int.MaxValue, minCol = int.MaxValue, maxRow = 0, maxCol = 0;

        foreach (var address in Cells.Keys)
        {
            hasAny = true;
            if (address.Row < minRow) { minRow = address.Row; }
            if (address.Row > maxRow) { maxRow = address.Row; }
            if (address.Column < minCol) { minCol = address.Column; }
            if (address.Column > maxCol) { maxCol = address.Column; }
        }

        foreach (var merged in MergedRanges)
        {
            hasAny = true;
            if (merged.Range.FirstRow < minRow) { minRow = merged.Range.FirstRow; }
            if (merged.Range.LastRow > maxRow) { maxRow = merged.Range.LastRow; }
            if (merged.Range.FirstColumn < minCol) { minCol = merged.Range.FirstColumn; }
            if (merged.Range.LastColumn > maxCol) { maxCol = merged.Range.LastColumn; }
        }

        return hasAny ? new CellRange(minRow, minCol, maxRow, maxCol) : null;
    }
}
