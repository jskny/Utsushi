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
        AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);
}
