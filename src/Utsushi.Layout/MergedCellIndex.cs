using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// セル→結合範囲の索引。<see cref="SheetModel.FindMergedRange"/>(結合範囲の線形探索)を
    /// セルごとに呼ぶと「セル数×結合範囲数」の計算量になるため、行方向を「かかる結合範囲の組が変わらない行の区間」に
    /// 分け、区間ごとにその区間にかかる結合範囲を列の先頭順に並べて持つ。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 区間の数は結合範囲の数の2倍+1以下であり、メモリは「区間数×1区間にかかる結合範囲の数」で抑えられる
    /// (結合範囲の行数や格子の行数には比例しない。行ごとに結合範囲を登録すると、全行にわたる縦長の結合範囲が
    /// 多数あるシートで「格子の行数×結合範囲数」の要素を確保してしまう。layout-fidelity-reviewer指摘)。
    /// 検索する列(<c>Create</c> の <c>columns</c>)にかからない結合範囲は登録しない。
    /// </para>
    /// <para>
    /// 結果は <see cref="SheetModel.FindMergedRange"/> と同じになる(不正なファイルで結合範囲が重なっている場合も、
    /// <see cref="SheetModel.MergedRanges"/> の並びで先に現れる範囲を返す)。
    /// </para>
    /// </remarks>
    internal sealed class MergedCellIndex
    {
        /// <summary>各区間の先頭行(昇順)。区間 i は <c>_segmentStarts[i]</c> から次の区間の先頭行の手前まで。</summary>
        private readonly int[] _segmentStarts;

        /// <summary>各区間にかかる結合範囲(列の先頭順、同じ先頭列なら元の並び順)。</summary>
        private readonly (int Order, MergedRange Range)[][] _segments;

        private MergedCellIndex(int[] segmentStarts, (int Order, MergedRange Range)[][] segments)
        {
            _segmentStarts = segmentStarts;
            _segments = segments;
        }

        /// <summary>
        /// 索引を作る。
        /// </summary>
        /// <param name="mergedRanges">シートの結合範囲(<see cref="SheetModel.MergedRanges"/>)。</param>
        /// <param name="columns">
        /// 検索する列番号。これらの列にかからない結合範囲は登録しない(null ならすべて登録する)。
        /// </param>
        public static MergedCellIndex Create(IReadOnlyList<MergedRange> mergedRanges, IReadOnlyList<int>? columns = null)
        {
            var minColumn = int.MinValue;
            var maxColumn = int.MaxValue;
            if (columns is { Count: > 0 })
            {
                minColumn = int.MaxValue;
                maxColumn = int.MinValue;
                foreach (var column in columns)
                {
                    minColumn = Math.Min(minColumn, column);
                    maxColumn = Math.Max(maxColumn, column);
                }
            }

            var entries = new List<(int Order, MergedRange Range)>();
            var boundaries = new SortedSet<int>();
            for (var order = 0; order < mergedRanges.Count; order++)
            {
                var merged = mergedRanges[order];
                var range = merged.Range;
                if (range.LastColumn < minColumn || range.FirstColumn > maxColumn)
                {
                    continue;
                }

                entries.Add((order, merged));
                boundaries.Add(range.FirstRow);
                boundaries.Add(range.LastRow + 1);
            }

            // 列の先頭で並べておくと、検索時に「先頭列がセルより右」の区間で打ち切れる。同じ先頭列どうしは元の並び順を保つ。
            entries.Sort((a, b) =>
            {
                var byColumn = a.Range.Range.FirstColumn.CompareTo(b.Range.Range.FirstColumn);
                return byColumn != 0 ? byColumn : a.Order.CompareTo(b.Order);
            });

            var segmentStarts = new int[boundaries.Count];
            boundaries.CopyTo(segmentStarts);
            var segments = new (int Order, MergedRange Range)[segmentStarts.Length][];
            var active = new List<(int Order, MergedRange Range)>();
            for (var i = 0; i < segmentStarts.Length; i++)
            {
                var row = segmentStarts[i];
                active.Clear();
                foreach (var entry in entries)
                {
                    if (entry.Range.Range.FirstRow <= row && row <= entry.Range.Range.LastRow)
                    {
                        active.Add(entry);
                    }
                }

                segments[i] = active.ToArray();
            }

            return new MergedCellIndex(segmentStarts, segments);
        }

        /// <summary>
        /// 指定セルを含む結合範囲を返す。結合されていない場合は null。
        /// </summary>
        public MergedRange? Find(CellAddress address)
        {
            var segment = UpperBound(_segmentStarts, address.Row) - 1;
            if (segment < 0)
            {
                return null;
            }

            MergedRange? found = null;
            var foundOrder = int.MaxValue;
            foreach (var (order, merged) in _segments[segment])
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

        /// <summary><paramref name="value"/> より大きい最初の要素の位置(二分探索)。</summary>
        private static int UpperBound(int[] sorted, int value)
        {
            var lo = 0;
            var hi = sorted.Length;
            while (lo < hi)
            {
                var mid = lo + ((hi - lo) / 2);
                if (sorted[mid] <= value)
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
