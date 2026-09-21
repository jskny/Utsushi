using SkiaSharp;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering
{
    /// <summary>
    /// 対応済み接続線プリセット(要件10.9補足)から、塗りつぶし無しの開いた <see cref="SKPath"/> を
    /// 組み立てる。接続線は塗りつぶし・調整ガイドを持たないため、<see cref="ShapeGeometryBuilder"/>
    /// とは別に新設する(design.md「Rendering レイヤー」参照)。
    /// </summary>
    /// <remarks>
    /// 接続点(コネクションサイト)の解決は行わず、アンカー矩形と反転フラグのみから経路を決める
    /// (要件10.9補足)。
    /// </remarks>
    internal static class ConnectorGeometryBuilder
    {
        public static SKPath Build(ConnectorPresetType preset, bool flipHorizontal, bool flipVertical, SKRect rect)
        {
            var (start, end) = ResolveEndpoints(rect, flipHorizontal, flipVertical);

            return preset switch
            {
                ConnectorPresetType.Straight => StraightPath(start, end),
                ConnectorPresetType.Bent2Segment => BentPath(start, new SKPoint(end.X, start.Y), end),
                ConnectorPresetType.Bent3Segment => Bent3Path(start, end),
                ConnectorPresetType.Curved2Segment => CurvedPath(start, new SKPoint(end.X, start.Y), end),
                ConnectorPresetType.Curved3Segment => Curved3Path(start, end),
                _ => StraightPath(start, end),
            };
        }

        /// <summary>
        /// 反転フラグに応じた始点・終点を求める。<paramref name="flipHorizontal"/>/
        /// <paramref name="flipVertical"/>が無い場合は矩形の左上→右下を結ぶ。
        /// </summary>
        private static (SKPoint Start, SKPoint End) ResolveEndpoints(SKRect rect, bool flipHorizontal, bool flipVertical)
        {
            var left = flipHorizontal ? rect.Right : rect.Left;
            var right = flipHorizontal ? rect.Left : rect.Right;
            var top = flipVertical ? rect.Bottom : rect.Top;
            var bottom = flipVertical ? rect.Top : rect.Bottom;
            return (new SKPoint(left, top), new SKPoint(right, bottom));
        }

        private static SKPath StraightPath(SKPoint start, SKPoint end)
        {
            var path = new SKPath();
            path.MoveTo(start);
            path.LineTo(end);
            return path;
        }

        /// <summary>水平→垂直(または反転により垂直→水平)の順で直角に1回折れる2辺。</summary>
        private static SKPath BentPath(SKPoint start, SKPoint bend, SKPoint end)
        {
            var path = new SKPath();
            path.MoveTo(start);
            path.LineTo(bend);
            path.LineTo(end);
            return path;
        }

        /// <summary>始点・終点のX中央で折り返す3辺(水平→垂直→水平)。</summary>
        private static SKPath Bent3Path(SKPoint start, SKPoint end)
        {
            var midX = (start.X + end.X) / 2f;
            var bend1 = new SKPoint(midX, start.Y);
            var bend2 = new SKPoint(midX, end.Y);

            var path = new SKPath();
            path.MoveTo(start);
            path.LineTo(bend1);
            path.LineTo(bend2);
            path.LineTo(end);
            return path;
        }

        /// <summary>始点・終点を結ぶ、<paramref name="control"/>を制御点とする2次ベジェ曲線1本。</summary>
        private static SKPath CurvedPath(SKPoint start, SKPoint control, SKPoint end)
        {
            var path = new SKPath();
            path.MoveTo(start);
            path.QuadTo(control, end);
            return path;
        }

        /// <summary><see cref="Bent3Path"/>と同じ折れ点を通る2本のベジェ曲線によるS字カーブ。</summary>
        private static SKPath Curved3Path(SKPoint start, SKPoint end)
        {
            var midX = (start.X + end.X) / 2f;
            var midY = (start.Y + end.Y) / 2f;
            var control1 = new SKPoint(midX, start.Y);
            var control2 = new SKPoint(midX, end.Y);

            var path = new SKPath();
            path.MoveTo(start);
            path.QuadTo(control1, new SKPoint(midX, midY));
            path.QuadTo(control2, end);
            return path;
        }
    }
}
