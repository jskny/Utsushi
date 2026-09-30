using System;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// 印刷範囲が設定されていないシートの印刷対象範囲(使用範囲)を求める(要件3.10)。
    /// </summary>
    /// <remarks>
    /// Excelは印刷範囲が無い場合、セルに加えて画像・図形などの描画オブジェクトが置かれた範囲も印刷する。
    /// セルだけの使用範囲(<see cref="SheetModel.GetUsedRange"/>)では、表の右や下に置いた図形や、
    /// 図形だけのシートが出力されないため、描画オブジェクトが占めるセルの範囲を合わせる。
    /// </remarks>
    internal static class UsedRangeResolver
    {
        /// <summary>
        /// 固定サイズ(<see cref="FixedAnchorExtent"/>)の描画オブジェクトの終端のセルを探すときに、
        /// 行・列をたどる回数の上限。描画オブジェクトの寸法は Layout の <c>MaxDrawingObjectDimensionPt</c> 相当に
        /// 収まるため通常は数十回で終わるが、幅・高さが0の(非表示の)行・列が続く入力でも打ち切れるようにする。
        /// </summary>
        private const int MaxWalkSteps = 4096;

        /// <summary>セルと描画オブジェクトの両方を含む使用範囲。どちらも無ければ null。</summary>
        public static CellRange? Resolve(SheetModel sheet, double maxDigitWidthPx)
        {
            var range = sheet.GetUsedRange();

            foreach (var drawingObject in sheet.DrawingObjects)
            {
                var objectRange = GetOccupiedRange(sheet, drawingObject, maxDigitWidthPx);
                range = range is { } current ? Union(current, objectRange) : objectRange;
            }

            return range;
        }

        /// <summary>描画オブジェクトが占めるセルの範囲(左上のアンカーセルから、終端が掛かるセルまで)。</summary>
        internal static CellRange GetOccupiedRange(SheetModel sheet, DrawingObjectModel drawingObject, double maxDigitWidthPx)
        {
            var from = drawingObject.AnchorCell;
            int lastRow, lastColumn;

            switch (drawingObject.Extent)
            {
                case CellSpanAnchorExtent span:
                    // 終端がセルの境界ちょうど(オフセット0)なら、そのセルには掛かっていない。
                    lastColumn = span.ToOffset.X > 0 ? span.ToCell.Column : span.ToCell.Column - 1;
                    lastRow = span.ToOffset.Y > 0 ? span.ToCell.Row : span.ToCell.Row - 1;
                    break;

                case FixedAnchorExtent size:
                    lastColumn = WalkColumns(sheet, from.Column, drawingObject.AnchorOffset.X + size.WidthPt, maxDigitWidthPx);
                    lastRow = WalkRows(sheet, from.Row, drawingObject.AnchorOffset.Y + size.HeightPt);
                    break;

                default:
                    lastColumn = from.Column;
                    lastRow = from.Row;
                    break;
            }

            lastColumn = Math.Max(from.Column, Math.Min(lastColumn, CellAddress.MaxColumn));
            lastRow = Math.Max(from.Row, Math.Min(lastRow, CellAddress.MaxRow));
            return new CellRange(from.Row, from.Column, lastRow, lastColumn);
        }

        private static int WalkColumns(SheetModel sheet, int column, double extentPt, double maxDigitWidthPx)
        {
            var remaining = extentPt;
            for (var step = 0; step < MaxWalkSteps && column < CellAddress.MaxColumn; step++)
            {
                var widthPt = ExcelUnitConverter.ColumnWidthToPoints(sheet.GetColumnWidth(column), maxDigitWidthPx);
                if (remaining <= widthPt)
                {
                    break;
                }

                remaining -= widthPt;
                column++;
            }

            return column;
        }

        private static int WalkRows(SheetModel sheet, int row, double extentPt)
        {
            var remaining = extentPt;
            for (var step = 0; step < MaxWalkSteps && row < CellAddress.MaxRow; step++)
            {
                var heightPt = sheet.GetRowHeight(row);
                if (remaining <= heightPt)
                {
                    break;
                }

                remaining -= heightPt;
                row++;
            }

            return row;
        }

        private static CellRange Union(CellRange a, CellRange b) =>
            new(
                Math.Min(a.FirstRow, b.FirstRow),
                Math.Min(a.FirstColumn, b.FirstColumn),
                Math.Max(a.LastRow, b.LastRow),
                Math.Max(a.LastColumn, b.LastColumn));
    }
}
