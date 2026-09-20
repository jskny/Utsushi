using System;
using System.Collections.Generic;

namespace Utsushi.Layout
{
    /// <summary>
    /// 行(または列)の並びを、手動改ページと印字可能領域に基づいてページ単位の帯へ分割する。
    /// </summary>
    /// <remarks>
    /// 要件3.2(手動改ページ)と要件3.3(自動改ページ)の両方をここで扱う。
    /// 行にも列にも同じロジックが適用されるため、汎用の「指標→サイズ」関数として実装する。
    /// </remarks>
    internal static class PageBandCalculator
    {
        /// <summary>
        /// 分割を行う。
        /// </summary>
        /// <param name="indices">分割対象の行番号または列番号(昇順)。</param>
        /// <param name="sizeOf">各指標のサイズ(ポイント)を返す関数。</param>
        /// <param name="availableSizePt">1ページの印字可能領域(ポイント、論理座標)。</param>
        /// <param name="manualBreaks">手動改ページの位置。その指標の「手前」で改ページする。</param>
        /// <returns>ページ単位に分割された指標の帯。常に1つ以上の帯を返す(入力が空の場合は空の帯1つ)。</returns>
        public static IReadOnlyList<IReadOnlyList<int>> Split(
            IReadOnlyList<int> indices,
            Func<int, double> sizeOf,
            double availableSizePt,
            IReadOnlyCollection<int> manualBreaks)
        {
            var bands = new List<IReadOnlyList<int>>();
            if (indices.Count == 0)
            {
                bands.Add(Array.Empty<int>());
                return bands;
            }

            var breakSet = manualBreaks as ISet<int> ?? new HashSet<int>(manualBreaks);
            var current = new List<int>();
            var currentSize = 0.0;

            foreach (var index in indices)
            {
                var size = sizeOf(index);

                var forcedBreak = breakSet.Contains(index);
                var overflows = currentSize + size > availableSizePt;

                // 帯が空のときは改ページしない。1つの行/列だけで印字可能領域を超える場合は
                // その行/列単独のページとし、はみ出し分はページ端で切り取られる(Excel と同じ挙動)。
                if (current.Count > 0 && (forcedBreak || overflows))
                {
                    bands.Add(current);
                    current = new List<int>();
                    currentSize = 0.0;
                }

                current.Add(index);
                currentSize += size;
            }

            if (current.Count > 0)
            {
                bands.Add(current);
            }

            return bands;
        }
    }
}
