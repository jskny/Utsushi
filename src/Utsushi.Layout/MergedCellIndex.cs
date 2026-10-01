using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// セル→結合範囲の索引。<see cref="SheetModel.FindMergedRange"/>(結合範囲の線形探索)を
    /// セルごとに呼ぶと「セル数×結合範囲数」の計算量になるため、格子の行ごとに、その行にかかる結合範囲の
    /// 列の区間を並べて持つ。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 登録するのは格子(<see cref="SheetGrid.Rows"/>)に含まれる行だけであり、メモリは
    /// 「結合範囲がかかる格子の行数」に比例する(結合範囲の面積には比例しない)。格子の行数は
    /// 印刷範囲のセル数の上限(<see cref="ReportLayoutEngine.MaxPrintRangeCells"/>)で抑えられている。
    /// </para>
    /// <para>
    /// 結果は <see cref="SheetModel.FindMergedRange"/> と同じになる(不正なファイルで結合範囲が重なっている場合も、
    /// <see cref="SheetModel.MergedRanges"/> の並びで先に現れる範囲を返す)。
    /// </para>
    /// </remarks>
    internal sealed class MergedCellIndex
    {
        private readonly Dictionary<int, List<(int Order, MergedRange Range)>> _byRow;

        private MergedCellIndex(Dictionary<int, List<(int Order, MergedRange Range)>> byRow)
        {
            _byRow = byRow;
        }

        /// <summary>
        /// 索引を作る。
        /// </summary>
        /// <param name="mergedRanges">シートの結合範囲(<see cref="SheetModel.MergedRanges"/>)。</param>
        /// <param name="rows">索引に登録する行番号(昇順)。</param>
        public static MergedCellIndex Create(IReadOnlyList<MergedRange> mergedRanges, IReadOnlyList<int> rows)
        {
            var byRow = new Dictionary<int, List<(int Order, MergedRange Range)>>();

            for (var order = 0; order < mergedRanges.Count; order++)
            {
                var merged = mergedRanges[order];
                var range = merged.Range;

                for (var i = LowerBound(rows, range.FirstRow); i < rows.Count && rows[i] <= range.LastRow; i++)
                {
                    if (!byRow.TryGetValue(rows[i], out var list))
                    {
                        list = new List<(int Order, MergedRange Range)>();
                        byRow[rows[i]] = list;
                    }

                    list.Add((order, merged));
                }
            }

            // 列の先頭で並べておくと、検索時に「先頭列がセルより右」の区間で打ち切れる。
            // 同じ先頭列の区間どうしは元の並び順を保つ。
            foreach (var list in byRow.Values)
            {
                list.Sort((a, b) =>
                {
                    var byColumn = a.Range.Range.FirstColumn.CompareTo(b.Range.Range.FirstColumn);
                    return byColumn != 0 ? byColumn : a.Order.CompareTo(b.Order);
                });
            }

            return new MergedCellIndex(byRow);
        }

        /// <summary>
        /// 指定セルを含む結合範囲を返す。結合されていない場合(または索引に登録していない行の場合)は null。
        /// </summary>
        public MergedRange? Find(CellAddress address)
        {
            if (!_byRow.TryGetValue(address.Row, out var list))
            {
                return null;
            }

            MergedRange? found = null;
            var foundOrder = int.MaxValue;
            foreach (var (order, merged) in list)
            {
                var range = merged.Range;
                if (range.FirstColumn > address.Column)
                {
                    break;
                }

                if (address.Column <= range.LastColumn && order < foundOrder)
                {
                    found = merged;
                    foundOrder = order;
                }
            }

            return found;
        }

        /// <summary><paramref name="value"/> 以上の最初の要素の位置(二分探索)。</summary>
        private static int LowerBound(IReadOnlyList<int> sorted, int value)
        {
            var lo = 0;
            var hi = sorted.Count;
            while (lo < hi)
            {
                var mid = lo + ((hi - lo) / 2);
                if (sorted[mid] < value)
                {
                    lo = mid + 1;
                }
                else
                {
                    hi = mid;
                }
            }

            return lo;
        }
    }
}
