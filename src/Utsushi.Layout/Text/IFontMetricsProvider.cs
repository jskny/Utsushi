using System;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout.Text
{
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
    public readonly struct FontMetrics : IEquatable<FontMetrics>
    {
        public FontMetrics(double ascentPt, double descentPt, double lineSpacingPt)
        {
            AscentPt = ascentPt;
            DescentPt = descentPt;
            LineSpacingPt = lineSpacingPt;
        }

        /// <summary>ベースラインから上端までの距離。</summary>
        public double AscentPt { get; }

        /// <summary>ベースラインから下端までの距離。</summary>
        public double DescentPt { get; }

        /// <summary>行送り(1行の高さ)。</summary>
        public double LineSpacingPt { get; }

        public override string ToString() =>
            $"FontMetrics {{ AscentPt = {AscentPt}, DescentPt = {DescentPt}, LineSpacingPt = {LineSpacingPt} }}";

        public bool Equals(FontMetrics other) =>
            AscentPt.Equals(other.AscentPt) && DescentPt.Equals(other.DescentPt) && LineSpacingPt.Equals(other.LineSpacingPt);

        public override bool Equals(object? obj) => obj is FontMetrics other && Equals(other);

        public override int GetHashCode() => HashCode.Combine(AscentPt, DescentPt, LineSpacingPt);

        public static bool operator ==(FontMetrics left, FontMetrics right) => left.Equals(right);

        public static bool operator !=(FontMetrics left, FontMetrics right) => !left.Equals(right);
    }
}
