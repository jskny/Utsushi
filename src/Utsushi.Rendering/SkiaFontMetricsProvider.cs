using System;
using System.Collections.Concurrent;
using SkiaSharp;
using Utsushi.Rendering.Fonts;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering
{
    /// <summary>
    /// SkiaSharp の実フォントメトリクスに基づく <see cref="IFontMetricsProvider"/> 実装。
    /// </summary>
    /// <remarks>
    /// Layout レイヤーはこのインターフェース越しにのみフォント情報へアクセスするため、
    /// SkiaSharp への依存は Rendering レイヤーに閉じる(`.kiro/steering/structure.md`)。
    /// </remarks>
    public sealed class SkiaFontMetricsProvider : IFontMetricsProvider
    {
        /// <summary>斜体を合成するときの傾き。一般的な斜体の角度(約12度)に合わせる。</summary>
        internal const float ItalicSkew = -0.21f;

        /// <summary><see cref="GetMetrics"/> の結果を記録する数の上限。超えたら記録を捨てて作り直す。</summary>
        internal const int MaxCachedMetrics = 1024;

        private readonly FontResolver _fontResolver;

        /// <summary>
        /// <see cref="GetMetrics"/> の結果(書体・サイズ・斜体の合成の有無ごと)。複数スレッドから同時に変換しても壊れないよう、
        /// スレッドセーフな辞書にする。
        /// </summary>
        private readonly ConcurrentDictionary<(SKTypeface Typeface, double SizePt, bool SynthesizeItalic), FontMetrics> _metricsCache = new();

        public SkiaFontMetricsProvider(FontResolver fontResolver)
        {
            _fontResolver = fontResolver ?? throw new ArgumentNullException(nameof(fontResolver));
            Shaper = new GlyphShaper(fontResolver);
        }

        /// <inheritdoc />
        /// <remarks>
        /// 結果は書体・サイズ・斜体の合成の有無ごとに記録して使い回す(レイアウト計算ではセルごとに呼ばれるため)。
        /// 書体の解決(<see cref="FontResolver.Resolve"/>)は毎回行うため、厳格モードでフォントが見つからない場合の例外は
        /// 従来どおり呼び出しのたびに発生する。
        /// </remarks>
        public FontMetrics GetMetrics(FontStyle font)
        {
            var resolved = _fontResolver.Resolve(font);
            var key = (resolved.Typeface, font.SizePt, resolved.SynthesizeItalic);
            if (_metricsCache.TryGetValue(key, out var cached))
            {
                return cached;
            }

            using var skFont = GlyphShaper.CreateFont(resolved, font.SizePt);
            var metrics = skFont.Metrics;

            // Skia は Ascent を負値、Descent を正値で返す。Utsushi 側は両方とも正の距離として扱う。
            var ascent = -metrics.Ascent;
            var descent = metrics.Descent;
            var lineSpacing = ascent + descent + metrics.Leading;

            var result = new FontMetrics(ascent, descent, lineSpacing);

            // 縮小表示(Excel の「縮小して全体を表示する」)ではサイズが連続的に変わるため、記録する数に上限を設ける。
            if (_metricsCache.Count >= MaxCachedMetrics)
            {
                _metricsCache.Clear();
            }

            _metricsCache[key] = result;
            return result;
        }

        /// <inheritdoc />
        /// <remarks>
        /// 外字用の代替フォントで描く文字(要件11.2)・異体字(要件11.3)も含め、描画と同じ字形の並び
        /// (<see cref="GlyphShaper"/>)で送り幅を合計する。
        /// </remarks>
        public double MeasureTextWidth(FontStyle font, string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0.0;
            }

            return MeasureShaped(Shaper.Shape(font, text), font.SizePt);
        }

        /// <summary>計測と描画で共有する、文字列 → 字形の並びの変換。</summary>
        internal GlyphShaper Shaper { get; }

        /// <summary>字形の並びの送り幅(ポイント)を合計する。</summary>
        internal static double MeasureShaped(ShapedText shaped, double sizePt)
        {
            var width = 0.0;
            foreach (var run in shaped.Runs)
            {
                using var skFont = GlyphShaper.CreateFont(run.Face, sizePt);
                width += skFont.MeasureText(run.Glyphs);
            }

            return width;
        }
    }
}
