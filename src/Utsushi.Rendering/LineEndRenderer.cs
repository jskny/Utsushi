using System;
using SkiaSharp;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering
{
    /// <summary>
    /// 接続線・線吹き出しの引き出し線の端に矢印(<c>a:headEnd</c>/<c>a:tailEnd</c>)を描く(要件10.18)。
    /// </summary>
    /// <remarks>
    /// 矢印の大きさは線の太さの倍数で近似する(小=2倍、中=3倍、大=5倍)。細い線でも見えるよう、
    /// 倍数を掛ける元の太さは<see cref="MinBaseWidthPt"/>以上とする。線の端は矢印の下で短縮しない
    /// (太い線では、とがった矢印の先端の脇に線の角がわずかに見えることがある)。
    /// </remarks>
    internal static class LineEndRenderer
    {
        /// <summary>矢印の大きさを決める元の線の太さの最小値(ポイント)。</summary>
        private const float MinBaseWidthPt = 1.0f;

        /// <summary><c>stealth</c>の切れ込みの深さ(矢印の長さに対する比率)。</summary>
        private const float StealthNotchRatio = 0.4f;

        /// <summary>
        /// 開いた線(<paramref name="path"/>の最初の輪郭)の始点に <paramref name="headEnd"/>、終点に
        /// <paramref name="tailEnd"/> の矢印を描く。
        /// </summary>
        public static void DrawOnPath(
            SKCanvas canvas, SKPath path, LineEndStyle? headEnd, LineEndStyle? tailEnd, SKColor color, float lineWidth)
        {
            if (headEnd is null && tailEnd is null)
            {
                return;
            }

            using var measure = new SKPathMeasure(path, forceClosed: false);
            var length = measure.Length;
            if (length <= 0f)
            {
                return;
            }

            if (headEnd is not null && measure.GetPositionAndTangent(0f, out var startPoint, out var startTangent))
            {
                // 始点の矢印は線の外側(進行方向と逆)を向く。
                Draw(canvas, startPoint, new SKPoint(-startTangent.X, -startTangent.Y), headEnd, color, lineWidth);
            }

            if (tailEnd is not null && measure.GetPositionAndTangent(length, out var endPoint, out var endTangent))
            {
                Draw(canvas, endPoint, endTangent, tailEnd, color, lineWidth);
            }
        }

        /// <summary>
        /// 折れ線(<paramref name="points"/>)の始点・終点に矢印を描く。
        /// </summary>
        public static void DrawOnPolyline(
            SKCanvas canvas, SKPoint[] points, LineEndStyle? headEnd, LineEndStyle? tailEnd, SKColor color, float lineWidth)
        {
            if (points.Length < 2)
            {
                return;
            }

            // 端の点と重なる点は向きを決められないため、重ならない最初の点までさかのぼる。
            if (headEnd is not null)
            {
                var next = 1;
                while (next < points.Length - 1 && points[next] == points[0])
                {
                    next++;
                }

                Draw(canvas, points[0], points[0] - points[next], headEnd, color, lineWidth);
            }

            if (tailEnd is not null)
            {
                var last = points.Length - 1;
                var previous = last - 1;
                while (previous > 0 && points[previous] == points[last])
                {
                    previous--;
                }

                Draw(canvas, points[last], points[last] - points[previous], tailEnd, color, lineWidth);
            }
        }

        /// <summary>
        /// 先端 <paramref name="tip"/> に、<paramref name="direction"/>(線の外側を向く向き)の矢印を描く。
        /// </summary>
        public static void Draw(SKCanvas canvas, SKPoint tip, SKPoint direction, LineEndStyle style, SKColor color, float lineWidth)
        {
            var directionLength = direction.Length;
            if (directionLength <= float.Epsilon || float.IsNaN(directionLength))
            {
                return;
            }

            var d = new SKPoint(direction.X / directionLength, direction.Y / directionLength);
            var n = new SKPoint(-d.Y, d.X);

            var baseWidth = Math.Max(lineWidth, MinBaseWidthPt);
            var width = baseWidth * Factor(style.Width);
            var length = baseWidth * Factor(style.Length);
            var halfWidth = width / 2f;

            var baseCenter = tip - Scale(d, length);

            using var paint = new SKPaint
            {
                Color = color,
                IsAntialias = true,
                Style = style.Type == LineEndType.Arrow ? SKPaintStyle.Stroke : SKPaintStyle.Fill,
                StrokeWidth = lineWidth,

                // 開いた矢印(arrow)の先端が、角の結合(既定はマイター)で線の太さの倍ほど突き出さないようにする。
                StrokeJoin = SKStrokeJoin.Round,
                StrokeCap = SKStrokeCap.Round,
            };

            using var path = new SKPath();
            switch (style.Type)
            {
                case LineEndType.Triangle:
                    path.MoveTo(tip);
                    path.LineTo(baseCenter + Scale(n, halfWidth));
                    path.LineTo(baseCenter - Scale(n, halfWidth));
                    path.Close();
                    break;

                case LineEndType.Stealth:
                    path.MoveTo(tip);
                    path.LineTo(baseCenter + Scale(n, halfWidth));
                    path.LineTo(tip - Scale(d, length * (1f - StealthNotchRatio)));
                    path.LineTo(baseCenter - Scale(n, halfWidth));
                    path.Close();
                    break;

                case LineEndType.Arrow:
                    path.MoveTo(baseCenter + Scale(n, halfWidth));
                    path.LineTo(tip);
                    path.LineTo(baseCenter - Scale(n, halfWidth));
                    break;

                case LineEndType.Diamond:
                    path.MoveTo(tip + Scale(d, length / 2f));
                    path.LineTo(tip + Scale(n, halfWidth));
                    path.LineTo(tip - Scale(d, length / 2f));
                    path.LineTo(tip - Scale(n, halfWidth));
                    path.Close();
                    break;

                case LineEndType.Oval:
                    path.AddOval(new SKRect(tip.X - (length / 2f), tip.Y - halfWidth, tip.X + (length / 2f), tip.Y + halfWidth));
                    var angle = (float)(Math.Atan2(d.Y, d.X) * 180.0 / Math.PI);
                    path.Transform(SKMatrix.CreateRotationDegrees(angle, tip.X, tip.Y));
                    break;
            }

            canvas.DrawPath(path, paint);
        }

        private static float Factor(LineEndSize size) => size switch
        {
            LineEndSize.Small => 2f,
            LineEndSize.Large => 5f,
            _ => 3f,
        };

        private static SKPoint Scale(SKPoint point, float factor) => new(point.X * factor, point.Y * factor);
    }
}
