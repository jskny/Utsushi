using System;
using System.Collections.Generic;
using System.Linq;

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
        /// 「収まるかどうか」の比較に使う許容誤差(ポイント)。行高・列幅の合計や倍率での割り算で生じる
        /// 浮動小数点の誤差で、ちょうど収まる行/列が次のページへ送られないようにする。
        /// </summary>
        internal const double TolerancePt = 1e-6;

        /// <summary>
        /// 分割を行う(1ページに使える大きさがどのページも同じ場合)。
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
            IReadOnlyCollection<int> manualBreaks) =>
            Split(indices, sizeOf, _ => availableSizePt, manualBreaks);

        /// <summary>
        /// 分割を行う。1ページに使える大きさは、帯の先頭の指標によって変わってよい
        /// (印刷タイトルを付けるページだけタイトルの分だけ狭くなる、など)。
        /// </summary>
        /// <param name="indices">分割対象の行番号または列番号(昇順)。</param>
        /// <param name="sizeOf">各指標のサイズ(ポイント)を返す関数。</param>
        /// <param name="availableSizeFor">帯の先頭の指標を受け取り、その帯(ページ)で使える大きさを返す関数。</param>
        /// <param name="manualBreaks">
        /// 手動改ページの位置。その指標の「手前」で改ページする。改ページ位置の行/列そのものが
        /// 並びに無い(非表示・印刷タイトルなど)場合も、「直前の指標 &lt; 改ページ位置 &lt;= 現在の指標」を満たす
        /// 指標の手前で改ページする。
        /// </param>
        /// <returns>ページ単位に分割された指標の帯。常に1つ以上の帯を返す(入力が空の場合は空の帯1つ)。</returns>
        public static IReadOnlyList<IReadOnlyList<int>> Split(
            IReadOnlyList<int> indices,
            Func<int, double> sizeOf,
            Func<int, double> availableSizeFor,
            IReadOnlyCollection<int> manualBreaks)
        {
            var bands = new List<IReadOnlyList<int>>();
            if (indices.Count == 0)
            {
                bands.Add(Array.Empty<int>());
                return bands;
            }

            var sortedBreaks = manualBreaks.Count == 0
                ? Array.Empty<int>()
                : manualBreaks.Distinct().OrderBy(b => b).ToArray();
            var breakCursor = 0;

            var current = new List<int>();
            var currentSize = 0.0;
            var currentAvailable = 0.0;
            int? previous = null;

            foreach (var index in indices)
            {
                var size = sizeOf(index);

                // 直前の指標 < 改ページ位置 <= 現在の指標 を満たす改ページがあるか。
                var forcedBreak = false;
                while (breakCursor < sortedBreaks.Length && sortedBreaks[breakCursor] <= index)
                {
                    if (previous is { } prev && sortedBreaks[breakCursor] > prev)
                    {
                        forcedBreak = true;
                    }

                    breakCursor++;
                }

                var overflows = currentSize + size > currentAvailable + TolerancePt;

                // 帯が空のときは改ページしない。1つの行/列だけで印字可能領域を超える場合は
                // その行/列単独のページとし、はみ出し分はページ端で切り取られる(Excel と同じ挙動)。
                if (current.Count > 0 && (forcedBreak || overflows))
                {
                    bands.Add(current);
                    current = new List<int>();
                    currentSize = 0.0;
                }

                if (current.Count == 0)
                {
                    currentAvailable = availableSizeFor(index);
                }

                current.Add(index);
                currentSize += size;
                previous = index;
            }

            if (current.Count > 0)
            {
                bands.Add(current);
            }

            return bands;
        }
    }
}
