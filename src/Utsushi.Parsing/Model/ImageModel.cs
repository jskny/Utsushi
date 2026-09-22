using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// シートに埋め込まれた画像(要件9)。
    /// </summary>
    /// <param name="Id">
    /// <c>NonVisualDrawingProperties/@id</c>。接続線の接続先解決(要件10.11)のために保持する。
    /// </param>
    /// <param name="Data">画像のバイナリ。</param>
    /// <param name="ContentType">MIMEタイプ(例: <c>"image/png"</c>)。</param>
    /// <param name="RotationDegrees"><c>a:xfrm/@rot</c> から変換した回転角(度、時計回り)。</param>
    /// <param name="AnchorCell">アンカー左上セル。</param>
    /// <param name="AnchorOffset">アンカーセル左上からのオフセット(ポイント)。</param>
    /// <param name="Extent">画像の終端(サイズ)の決め方。</param>
    public sealed record ImageModel(
        uint Id,
        byte[] Data,
        string ContentType,
        double RotationDegrees,
        CellAddress AnchorCell,
        PointPt AnchorOffset,
        AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);
}
