using System;
using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// 1シート分の内部モデル。OOXML の型を一切含まない POCO。
    /// </summary>
    /// <param name="Name">シート名。</param>
    /// <param name="Cells">値または書式を持つセル。空セルは含まれない場合がある。</param>
    /// <param name="MergedRanges">結合セル範囲。</param>
    /// <param name="ColumnWidths">列幅。索引0が列A。単位はExcelの「文字数」単位。</param>
    /// <param name="RowHeights">行高。索引0が行1。単位はポイント。</param>
    /// <param name="DefaultColumnWidth">
    /// <paramref name="ColumnWidths"/> の範囲外の列に適用する既定列幅(文字数単位)。<c>sheetFormatPr/@defaultColWidth</c>
    /// が無い場合は <see cref="double.NaN"/> で、既定列幅は <paramref name="BaseColumnWidth"/> と最大数字幅から
    /// Layoutレイヤーが求める(<paramref name="ColumnWidths"/> の要素も、幅の指定が無い列は NaN になる)。
    /// </param>
    /// <param name="DefaultRowHeight"><paramref name="RowHeights"/> の範囲外の行に適用する既定行高(ポイント)。</param>
    /// <param name="HiddenColumns">非表示の列番号(1始まり)。</param>
    /// <param name="HiddenRows">非表示の行番号(1始まり)。</param>
    /// <param name="PageSetup">ページ設定。</param>
    /// <param name="DrawingObjects">
    /// シート上の画像・図形(要件9, 10)。<c>drawing.xml</c> の出現順を保持する
    /// (Excelは画像・図形をこの順で重ねて描画するため)。
    /// </param>
    /// <param name="BaseColumnWidth"><c>sheetFormatPr/@baseColWidth</c>(文字数。既定8)。既定列幅の算出に使う。</param>
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
        PageSetupModel PageSetup,
        IReadOnlyList<DrawingObjectModel> DrawingObjects,
        int BaseColumnWidth = 8)
    {
        private static readonly IReadOnlyDictionary<int, CellStyle> NoRowStyles = new Dictionary<int, CellStyle>();

        /// <summary>
        /// ブックの標準の書式(<c>cellXfs</c> の0番)。ファイルにセルも行・列の書式も無い位置に使う(要件1.10)。
        /// </summary>
        public CellStyle DefaultCellStyle { get; init; } = CellStyle.Default;

        /// <summary>行の書式(<c>row/@s</c>。<c>@customFormat</c> が真の行だけ)。キーは1始まりの行番号(要件1.10)。</summary>
        public IReadOnlyDictionary<int, CellStyle> RowStyles { get; init; } = NoRowStyles;

        /// <summary>
        /// 列の書式(<c>col/@style</c>)。列番号の昇順に並び、範囲は重ならない(要件1.10)。
        /// 16,384 列までの <c>col</c> 要素を列ごとに展開せず、範囲のまま持つ。
        /// </summary>
        public IReadOnlyList<ColumnStyleRange> ColumnStyles { get; init; } = Array.Empty<ColumnStyleRange>();

        /// <summary>
        /// 指定位置に適用される書式を返す。セルがあればセルの書式、無ければ行の書式、列の書式、
        /// ブックの標準の書式の順に最初に見つかったもの(Excel と同じ。要件1.10)。
        /// </summary>
        public CellStyle GetEffectiveStyle(CellAddress address)
        {
            if (Cells.TryGetValue(address, out var cell))
            {
                return cell.Style;
            }

            if (RowStyles.TryGetValue(address.Row, out var rowStyle))
            {
                return rowStyle;
            }

            return FindColumnStyle(address.Column) ?? DefaultCellStyle;
        }

        private CellStyle? FindColumnStyle(int column)
        {
            var low = 0;
            var high = ColumnStyles.Count - 1;
            while (low <= high)
            {
                var mid = low + ((high - low) / 2);
                var range = ColumnStyles[mid];
                if (column < range.FirstColumn)
                {
                    high = mid - 1;
                }
                else if (column > range.LastColumn)
                {
                    low = mid + 1;
                }
                else
                {
                    return range.Style;
                }
            }

            return null;
        }

        /// <summary>指定セルを取得する。存在しない場合は null。</summary>
        public CellModel? GetCell(CellAddress address) =>
            Cells.TryGetValue(address, out var cell) ? cell : null;

        /// <summary>
        /// 1始まりの列番号に対する列幅(文字数単位)。定義が無い列には既定値を返す。既定列幅が暗黙
        /// (<c>defaultColWidth</c> 無し)の場合は <see cref="double.NaN"/> を返す(<see cref="DefaultColumnWidth"/> 参照)。
        /// </summary>
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

        /// <summary>
        /// 指定セルが結合範囲の内側にあり、かつ先頭(アンカー)セルではないかどうかを判定する。
        /// </summary>
        /// <param name="address">判定対象のセル。</param>
        /// <param name="merged">非アンカー位置だった場合、そのセルが属する結合範囲。それ以外は null。</param>
        public bool IsNonAnchorMergedCell(CellAddress address, out MergedRange? merged)
        {
            var found = FindMergedRange(address);
            if (found is not null && found.Anchor != address)
            {
                merged = found;
                return true;
            }

            merged = null;
            return false;
        }

        /// <summary>
        /// セルが存在する範囲(使用範囲)。セルが1つも無い場合は null。
        /// 行・列の書式(<see cref="RowStyles"/>・<see cref="ColumnStyles"/>)は含めない(要件1.10補足)。
        /// </summary>
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

    /// <summary>列の書式(<c>col/@style</c>)を持つ列の範囲(要件1.10)。</summary>
    /// <param name="FirstColumn">先頭の列番号(1始まり)。</param>
    /// <param name="LastColumn">末尾の列番号(1始まり、両端を含む)。</param>
    /// <param name="Style">書式。</param>
    public sealed record ColumnStyleRange(int FirstColumn, int LastColumn, CellStyle Style);
}
