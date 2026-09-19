using System;
using SkiaSharp;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering;

/// <summary>
/// SkiaSharp の実フォントメトリクスに基づく <see cref="IFontMetricsProvider"/> 実装。
/// </summary>
/// <remarks>
/// Layout レイヤーはこのインターフェース越しにのみフォント情報へアクセスするため、
/// SkiaSharp への依存は Rendering レイヤーに閉じる(`.kiro/steering/structure.md`)。
/// </remarks>
public sealed class SkiaFontMetricsProvider : IFontMetricsProvider
{
    private readonly FontResolver _fontResolver;

    public SkiaFontMetricsProvider(FontResolver fontResolver)
    {
        _fontResolver = fontResolver ?? throw new ArgumentNullException(nameof(fontResolver));
    }

    /// <inheritdoc />
    public FontMetrics GetMetrics(FontStyle font)
    {
        using var skFont = CreateFont(font);
        var metrics = skFont.Metrics;

        // Skia は Ascent を負値、Descent を正値で返す。Utsushi 側は両方とも正の距離として扱う。
        var ascent = -metrics.Ascent;
        var descent = metrics.Descent;
        var lineSpacing = ascent + descent + metrics.Leading;

        return new FontMetrics(ascent, descent, lineSpacing);
    }

    /// <inheritdoc />
    public double MeasureTextWidth(FontStyle font, string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return 0.0;
        }

        using var skFont = CreateFont(font);
        return MeasureText(skFont, text);
    }

    /// <summary>
    /// 文字列の送り幅(ポイント)を測る。
    /// </summary>
    /// <remarks>
    /// SkiaSharp 2.88 の <c>SKFont.MeasureText</c> はグリフ列のみを受け取るため、
    /// いったん文字列をグリフへ変換してから測る。
    /// <c>SKCanvas.DrawText(string, ...)</c> も同じ変換を行うため、描画結果と一致する。
    /// </remarks>
    internal static double MeasureText(SKFont font, string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return 0.0;
        }

        var glyphCount = font.CountGlyphs(text);
        if (glyphCount <= 0)
        {
            return 0.0;
        }

        var glyphs = new ushort[glyphCount];
        font.GetGlyphs(text, glyphs);
        return font.MeasureText(glyphs);
    }

    /// <summary>
    /// <see cref="FontStyle"/> から SkiaSharp のフォントを作る。
    /// </summary>
    /// <remarks>
    /// PDF の座標系は1単位=1ポイントであり、Skia の <c>TextSize</c> もその単位で解釈される。
    /// そのためフォントサイズ(pt)をそのまま渡してよい。
    /// </remarks>
    internal SKFont CreateFont(FontStyle font)
    {
        var typeface = _fontResolver.Resolve(font);
        return new SKFont(typeface, (float)font.SizePt)
        {
            Subpixel = true,
            Edging = SKFontEdging.Antialias,
        };
    }
}
