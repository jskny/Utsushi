using System;
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

        private readonly FontResolver _fontResolver;

        public SkiaFontMetricsProvider(FontResolver fontResolver)
        {
            _fontResolver = fontResolver ?? throw new ArgumentNullException(nameof(fontResolver));
            Shaper = new GlyphShaper(fontResolver);
        }

        /// <inheritdoc />
        public FontMetrics GetMetrics(FontStyle font)
        {
            using var skFont = CreateFont(font, out _);
            var metrics = skFont.Metrics;

            // Skia は Ascent を負値、Descent を正値で返す。Utsushi 側は両方とも正の距離として扱う。
            var ascent = -metrics.Ascent;
            var descent = metrics.Descent;
            var lineSpacing = ascent + descent + metrics.Leading;

            return new FontMetrics(ascent, descent, lineSpacing);
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

        /// <summary>
        /// <see cref="FontStyle"/> から SkiaSharp のフォントを作る。
        /// </summary>
        /// <remarks>
        /// <para>
        /// PDF の座標系は1単位=1ポイントであり、Skia の <c>TextSize</c> もその単位で解釈される。
        /// そのためフォントサイズ(pt)をそのまま渡してよい。
        /// </para>
        /// <para>
        /// 斜体の字形を持たないフォントには、ここで傾き(<see cref="SKFont.SkewX"/>)を与えて斜体を再現する。
        /// 太字は文字送り幅に影響させないため、描画時の輪郭の太らせで再現する
        /// (<paramref name="synthesizeBold"/> が true のとき)。
        /// </para>
        /// </remarks>
        /// <param name="font">帳票側のフォント指定。</param>
        /// <param name="synthesizeBold">
        /// 描画側で輪郭を太らせて太字を再現する必要があるかどうか。
        /// </param>
        internal SKFont CreateFont(FontStyle font, out bool synthesizeBold)
        {
            var resolved = _fontResolver.Resolve(font);
            synthesizeBold = resolved.SynthesizeBold;

            var skFont = new SKFont(resolved.Typeface, (float)font.SizePt)
            {
                Subpixel = true,
                Edging = SKFontEdging.Antialias,
            };

            if (resolved.SynthesizeItalic)
            {
                skFont.SkewX = ItalicSkew;
            }

            return skFont;
        }
    }
}
