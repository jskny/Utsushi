using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// シートに浮かぶ描画オブジェクト(画像・図形。要件9, 10)の共通の位置決め情報。
    /// </summary>
    /// <param name="AnchorCell">アンカー左上セル。</param>
    /// <param name="AnchorOffset">アンカーセル左上からのオフセット(ポイント)。</param>
    /// <param name="Extent">サイズ・終端の決め方。</param>
    public abstract record DrawingObjectModel(
        CellAddress AnchorCell,
        PointPt AnchorOffset,
        AnchorExtent Extent);

    /// <summary>描画オブジェクトのサイズ・終端の決め方。</summary>
    public abstract record AnchorExtent;

    /// <summary>
    /// OOXMLの <c>xdr:oneCellAnchor</c> 相当。セルに対して固定サイズを持ち、
    /// セルの拡大縮小に連動しない。
    /// </summary>
    public sealed record FixedAnchorExtent(double WidthPt, double HeightPt) : AnchorExtent;

    /// <summary>
    /// OOXMLの <c>xdr:twoCellAnchor</c> 相当。対角のセル+オフセットで範囲が決まり、
    /// セルの拡大縮小に連動する。
    /// </summary>
    public sealed record CellSpanAnchorExtent(CellAddress ToCell, PointPt ToOffset) : AnchorExtent;
}
