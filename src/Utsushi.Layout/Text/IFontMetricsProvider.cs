using Utsushi.Parsing.Model;

namespace Utsushi.Layout.Text;

/// <summary>
/// フォントのメトリクスと文字列幅を提供する。
/// </summary>
/// <remarks>
/// <para>
/// Layout レイヤーはテキストの配置(左右/上下揃え、縮小表示、折り返し)を計算するために
/// 実フォントのメトリクスを必要とするが、描画ライブラリ(SkiaSharp)に直接依存してはならない。
/// そのためインターフェースを Layout 側に置き、実装を Rendering 側に置く
/// (`.kiro/steering/structure.md` のレイヤー依存ルール)。
/// </para>
/// <para>単位はすべてポイント。</para>
/// </remarks>
public interface IFontMetricsProvider
{
    /// <summary>指定フォントのメトリクスを取得する。</summary>
    FontMetrics GetMetrics(FontStyle font);

    /// <summary>指定フォントで文字列を描画したときの送り幅(ポイント)を返す。</summary>
    double MeasureTextWidth(FontStyle font, string text);
}

/// <summary>
/// フォントメトリクス(ポイント単位、いずれも正の値)。
/// </summary>
/// <param name="AscentPt">ベースラインから上端までの距離。</param>
/// <param name="DescentPt">ベースラインから下端までの距離。</param>
/// <param name="LineSpacingPt">行送り(1行の高さ)。</param>
public readonly record struct FontMetrics(double AscentPt, double DescentPt, double LineSpacingPt);
