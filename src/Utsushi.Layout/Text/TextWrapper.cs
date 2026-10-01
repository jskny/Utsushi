using System;
using System.Collections.Generic;
using System.Globalization;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout.Text
{
    /// <summary>
    /// セル幅・図形の幅に合わせてテキストを折り返す(要件4.6)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 改行は LF・CRLF・CR 単独のいずれも段落区切りとして扱う。折り返し位置は書記素クラスタ
    /// (サロゲートペア・結合文字・異体字セレクタを含む、利用者が1文字と認識する単位)の境界に限る。
    /// UTF-16 の1単位ごとに区切ると、「𠮷」のようなサロゲートペアの途中で改行され、両側が文字化けする。
    /// </para>
    /// <para>
    /// 折り返しの規則は「行の先頭の1文字は必ず置き、以降は、行の文字列にその文字を足した幅
    /// (<see cref="IFontMetricsProvider.MeasureTextWidth"/>)が使える幅を超えたら、その文字の前で改行する」である。
    /// 素直に1文字足すごとに行全体を測り直すと、行の長さの2乗に比例する計測が必要になる。ここでは、
    /// 計測済みの幅から1文字あたりの平均の幅を求めて改行位置の見当を付け、その位置と1つ後ろの位置を
    /// 行全体の計測で確かめる(多くの行で計測は2回で済む)。改行するかどうかの判定には必ず行全体の計測値を使い、
    /// 見当は計測する位置を選ぶためだけに使う(文字ごとの幅の合計は、浮動小数点の丸めやカーニングのために
    /// 行全体の計測と一致するとは限らないため)。見当が外れ続けた場合は二分探索に切り替える。
    /// </para>
    /// <para>
    /// この探し方は、文字列の後ろに文字を足しても計測した幅が減らない(単調である)ことを前提にしており、
    /// その前提のもとで、1文字ずつ測り直す方法と同じ位置で改行する。字形の送り幅は負にならないため、
    /// 実フォントの計測(Rendering レイヤーの実装)でもこの前提は成り立つ。
    /// </para>
    /// </remarks>
    internal static class TextWrapper
    {
        /// <summary>テキストを折り返した行の並びを返す。空文字列の場合は空の1行を返す。</summary>
        public static List<string> Wrap(IFontMetricsProvider fontMetrics, FontStyle font, string text, double availableWidthPt)
        {
            var lines = new List<string>();
            var unified = text.IndexOf('\r') < 0 ? text : text.Replace("\r\n", "\n").Replace('\r', '\n');

            foreach (var paragraph in unified.Split('\n'))
            {
                if (paragraph.Length == 0)
                {
                    lines.Add(string.Empty);
                    continue;
                }

                new ParagraphWrapper(fontMetrics, font, paragraph, availableWidthPt).AppendLines(lines);
            }

            return lines.Count == 0 ? new List<string> { string.Empty } : lines;
        }

        /// <summary>1段落分の折り返し。</summary>
        private sealed class ParagraphWrapper
        {
            /// <summary>見当による絞り込みをこの回数まで試し、決まらなければ二分探索に切り替える。</summary>
            private const int MaxEstimateRounds = 2;

            private readonly IFontMetricsProvider _fontMetrics;
            private readonly FontStyle _font;
            private readonly string _paragraph;
            private readonly double _availableWidthPt;

            /// <summary>書記素クラスタの開始位置。末尾に段落の長さを番兵として置く(要素数 = クラスタ数 + 1)。</summary>
            private readonly List<int> _offsets;

            /// <summary>直近の計測から求めた、1クラスタあたりの平均の幅(見当用)。未計測なら NaN。</summary>
            private double _averageAdvance = double.NaN;

            public ParagraphWrapper(IFontMetricsProvider fontMetrics, FontStyle font, string paragraph, double availableWidthPt)
            {
                _fontMetrics = fontMetrics;
                _font = font;
                _paragraph = paragraph;
                _availableWidthPt = availableWidthPt;

                _offsets = new List<int>();
                var elements = StringInfo.GetTextElementEnumerator(paragraph);
                while (elements.MoveNext())
                {
                    _offsets.Add(elements.ElementIndex);
                }

                _offsets.Add(paragraph.Length);
            }

            private int ElementCount => _offsets.Count - 1;

            public void AppendLines(List<string> lines)
            {
                var start = 0;
                while (start < ElementCount)
                {
                    var end = FindLineEnd(start);
                    lines.Add(Substring(start, end));
                    start = end;
                }
            }

            /// <summary>
            /// <paramref name="start"/> 番目のクラスタから始まる行の終わり(次の行の先頭のクラスタの番号)を返す。
            /// すなわち、<paramref name="start"/> より後ろで、行の先頭からそのクラスタまでの幅が使える幅を超える
            /// 最初のクラスタの番号。超えるクラスタが無ければ段落の末尾(<see cref="ElementCount"/>)。
            /// </summary>
            /// <remarks>
            /// 探索中は、収まることが分かっている最後の位置 <c>fits</c>(行の先頭のクラスタは幅によらず置くため、
            /// 初期値は <paramref name="start"/>)と、超えることが分かっている最初の位置 <c>exceeds</c>
            /// (初期値は段落の末尾の番兵。番兵の位置は計測しない)の間を絞り込む。
            /// </remarks>
            private int FindLineEnd(int start)
            {
                var count = ElementCount;
                if (start + 1 >= count)
                {
                    return count;
                }

                var fits = start;
                var fitsWidth = double.NaN;
                var exceeds = count;
                var exceedsWidth = double.NaN;

                // 段落の最初の行では、段落全体が収まるかを先に確かめる(大半のセルは1回の計測で済む)。
                if (start == 0)
                {
                    var width = Measure(start, count - 1);
                    if (!(width > _availableWidthPt))
                    {
                        return count;
                    }

                    exceeds = count - 1;
                    exceedsWidth = width;
                }

                var rounds = 0;
                while (exceeds - fits > 1)
                {
                    var probe = rounds < MaxEstimateRounds
                        ? EstimateLastFitting(start, fits, fitsWidth, exceeds, exceedsWidth)
                        : -1;
                    rounds++;

                    if (probe < 0)
                    {
                        // 見当が付かない(付けても外れ続けた)場合は二分探索する。超える位置が未確定(番兵)の
                        // ときは、収まる側から間隔を倍々に広げる(行の長さに比べて段落が長い場合に、
                        // 段落の残り全体を何度も測らないため)。
                        probe = exceeds == count
                            ? Math.Min(count - 1, fits + Math.Max(1, fits - start))
                            : fits + ((exceeds - fits) / 2);
                        if (Exceeds(start, probe, out var width))
                        {
                            exceeds = probe;
                            exceedsWidth = width;
                        }
                        else
                        {
                            fits = probe;
                            fitsWidth = width;
                        }

                        continue;
                    }

                    // 見当の位置(収まると見込んだ最後の位置)と、その1つ後ろを確かめる。
                    if (Exceeds(start, probe, out var probeWidth))
                    {
                        exceeds = probe;
                        exceedsWidth = probeWidth;
                        if (probe - 1 > fits)
                        {
                            if (Exceeds(start, probe - 1, out var previousWidth))
                            {
                                exceeds = probe - 1;
                                exceedsWidth = previousWidth;
                            }
                            else
                            {
                                fits = probe - 1;
                                fitsWidth = previousWidth;
                            }
                        }
                    }
                    else
                    {
                        fits = probe;
                        fitsWidth = probeWidth;
                        if (probe + 1 < exceeds)
                        {
                            if (Exceeds(start, probe + 1, out var nextWidth))
                            {
                                exceeds = probe + 1;
                                exceedsWidth = nextWidth;
                            }
                            else
                            {
                                fits = probe + 1;
                                fitsWidth = nextWidth;
                            }
                        }
                    }
                }

                return exceeds;
            }

            /// <summary>
            /// 計測済みの幅から、行の先頭 <paramref name="start"/> から数えて収まる最後のクラスタの位置を見積もる。
            /// 戻り値は <paramref name="fits"/> より後ろ、<paramref name="exceeds"/> より前に収める。見積もれない場合は -1。
            /// </summary>
            private int EstimateLastFitting(int start, int fits, double fitsWidth, int exceeds, double exceedsWidth)
            {
                // 基準点: 収まる側の計測済みの位置。未計測なら「行の先頭の1つ手前で幅0」とみなす。
                var basePosition = double.IsNaN(fitsWidth) ? start - 1 : fits;
                var baseWidth = double.IsNaN(fitsWidth) ? 0.0 : fitsWidth;

                var average = !double.IsNaN(exceedsWidth)
                    ? (exceedsWidth - baseWidth) / (exceeds - basePosition)
                    : !double.IsNaN(fitsWidth) ? fitsWidth / (fits - start + 1) : _averageAdvance;

                if (!(average > 0) || double.IsInfinity(average))
                {
                    return -1;
                }

                var estimate = basePosition + Math.Floor((_availableWidthPt - baseWidth) / average);
                if (double.IsNaN(estimate))
                {
                    return -1;
                }

                return (int)Math.Max(fits + 1, Math.Min(exceeds - 1, estimate));
            }

            /// <summary>
            /// <paramref name="start"/> 番目から <paramref name="last"/> 番目(両端を含む)までのクラスタの幅が、
            /// 使える幅を超えるかどうか。
            /// </summary>
            private bool Exceeds(int start, int last, out double width)
            {
                width = Measure(start, last);
                return width > _availableWidthPt;
            }

            /// <summary>
            /// <paramref name="start"/> 番目から <paramref name="last"/> 番目(両端を含む)までのクラスタの幅を測り、
            /// 平均の幅(見当用)を更新する。
            /// </summary>
            private double Measure(int start, int last)
            {
                var width = _fontMetrics.MeasureTextWidth(_font, Substring(start, last + 1));
                if (width > 0 && !double.IsInfinity(width))
                {
                    _averageAdvance = width / (last - start + 1);
                }

                return width;
            }

            /// <summary><paramref name="first"/> 番目から <paramref name="end"/> 番目の手前までのクラスタの文字列。</summary>
            private string Substring(int first, int end) =>
                _paragraph.Substring(_offsets[first], _offsets[end] - _offsets[first]);
        }
    }
}
