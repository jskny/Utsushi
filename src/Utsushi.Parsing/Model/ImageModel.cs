using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// シートに埋め込まれた画像(要件9)。
    /// </summary>
    /// <param name="Data">画像のバイナリ。</param>
    /// <param name="ContentType">MIMEタイプ(例: <c>"image/png"</c>)。</param>
    /// <param name="AnchorCell">アンカー左上セル。</param>
    /// <param name="AnchorOffset">アンカーセル左上からのオフセット(ポイント)。</param>
    /// <param name="Extent">画像の終端(サイズ)の決め方。</param>
    public sealed record ImageModel(
        byte[] Data,
        string ContentType,
        CellAddress AnchorCell,
        PointPt AnchorOffset,
        ImageExtent Extent);

    /// <summary>画像のサイズ・終端の決め方。</summary>
    public abstract record ImageExtent;

    /// <summary>
    /// OOXMLの <c>xdr:oneCellAnchor</c> 相当。セルに対して固定サイズを持ち、
    /// セルの拡大縮小に連動しない。
    /// </summary>
    public sealed record FixedImageExtent(double WidthPt, double HeightPt) : ImageExtent;

    /// <summary>
    /// OOXMLの <c>xdr:twoCellAnchor</c> 相当。対角のセル+オフセットで範囲が決まり、
    /// セルの拡大縮小に連動する。
    /// </summary>
    public sealed record CellSpanImageExtent(CellAddress ToCell, PointPt ToOffset) : ImageExtent;
}
