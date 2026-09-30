using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// 接続線(<c>xdr:cxnSp</c>)の接続点(コネクションサイト)を、参照先図形のページ矩形上の
    /// 実際の座標として解決する(要件10.11)。
    /// </summary>
    /// <remarks>
    /// 既定では配置矩形の上下左右の中点(4方向。ECMA-376で<c>cxnLst</c>を持たない図形の
    /// 既定の接続点と同じ考え方)で近似する。<see cref="ShapePresetType.FlowChartInputOutput"/>
    /// (平行四辺形)のみ、左右の接続点(idx 1, 3)が実際の傾いた辺の中点からずれるため補正する
    /// (design.md参照。上下の辺は水平のため補正不要)。<see cref="ShapePresetType.FlowChartDocument"/>
    /// (波形)は、波形の谷の最深点が設計上ちょうど下辺中点と一致するため補正不要である。
    /// </remarks>
    internal static class ConnectionSiteResolver
    {
        /// <summary>
        /// <paramref name="rect"/>(参照先図形のページ矩形)上の接続点<paramref name="siteIndex"/>
        /// (4方向の範囲外の値は<c>% 4</c>で丸める)を解決する。<paramref name="preset"/>が
        /// <c>null</c>(画像・グループ、または非対応プリセット)の場合は既定の4方向近似を返す。
        /// </summary>
        /// <remarks>
        /// 反転した図形(要件10.17)では、反転前の図形上の接続点を配置矩形の中心を軸に鏡映した位置になる。
        /// </remarks>
        public static PointPt Resolve(
            RectPt rect, ShapePresetType? preset, uint siteIndex, bool flipHorizontal = false, bool flipVertical = false)
        {
            var point = ResolveUnflipped(rect, preset, siteIndex);
            return new PointPt(
                flipHorizontal ? rect.Left + rect.Right - point.X : point.X,
                flipVertical ? rect.Top + rect.Bottom - point.Y : point.Y);
        }

        private static PointPt ResolveUnflipped(RectPt rect, ShapePresetType? preset, uint siteIndex)
        {
            var index = (int)(siteIndex % 4);
            var midX = rect.Left + (rect.Width / 2.0);
            var midY = rect.Top + (rect.Height / 2.0);

            if (preset == ShapePresetType.FlowChartInputOutput && (index == 1 || index == 3))
            {
                var halfSkew = rect.Width * ShapeGeometryConstants.InputOutputSkewRatio / 2.0;
                var x = index == 1 ? rect.Left + halfSkew : rect.Right - halfSkew;
                return new PointPt(x, midY);
            }

            return index switch
            {
                0 => new PointPt(midX, rect.Top),
                1 => new PointPt(rect.Left, midY),
                2 => new PointPt(midX, rect.Bottom),
                _ => new PointPt(rect.Right, midY),
            };
        }
    }
}
