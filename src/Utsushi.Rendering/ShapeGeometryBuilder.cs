using System;
using System.Collections.Generic;
using SkiaSharp;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering
{
    /// <summary>
    /// 対応済みプリセットジオメトリ(要件10.1補足)から <see cref="SKPath"/> を組み立てる。
    /// </summary>
    /// <remarks>
    /// 調整ガイド値(<see cref="ShapeModel.AdjustmentValues"/>)が無い(<see cref="double.NaN"/>)位置は
    /// ここで定義するECMA-376の既定値を補う。既定値・角度単位はこの実装時点でECMA-376の一次資料へ
    /// 当たり直せていないため、実際にExcelで生成した図形との見た目の突き合わせは別途行う
    /// (design.md「Rendering レイヤー」参照)。
    /// </remarks>
    internal static class ShapeGeometryBuilder
    {
        private const double DefaultRoundRectAdj = 0.16667;
        private const double DefaultTriangleApexAdj = 0.5;
        private const double DefaultArrowShaftAdj = 0.5;
        private const double DefaultArrowHeadAdj = 0.5;

        /// <summary>
        /// 双方向矢印(leftRightArrow/upDownArrow)の矢尻長さ比の既定値。単方向矢印と同じ0.5を使うと
        /// 両端の矢尻が幅の合計100%を占めて軸(シャフト)が消え、菱形に潰れてしまうため、
        /// 単方向より小さい既定値にする(片方向矢印とは既定値の意味が異なる近似実装)。
        /// </summary>
        private const double DefaultDoubleArrowHeadAdj = 0.25;

        private const double DefaultCalloutTipXAdj = -0.25;
        private const double DefaultCalloutTipYAdj = 0.75;

        public static SKPath Build(ShapePresetType preset, IReadOnlyList<double> adjustmentValues, SKRect rect) =>
            preset switch
            {
                ShapePresetType.Rect => RectPath(rect),
                ShapePresetType.Ellipse => EllipsePath(rect),
                ShapePresetType.RoundRect => RoundRectPath(rect, Adj(adjustmentValues, 0, DefaultRoundRectAdj)),
                ShapePresetType.Triangle => TrianglePath(rect, Adj(adjustmentValues, 0, DefaultTriangleApexAdj)),
                ShapePresetType.RightArrow => ArrowPath(rect, ArrowDirection.Right, ShaftAdj(adjustmentValues), HeadAdj(adjustmentValues)),
                ShapePresetType.LeftArrow => ArrowPath(rect, ArrowDirection.Left, ShaftAdj(adjustmentValues), HeadAdj(adjustmentValues)),
                ShapePresetType.UpArrow => ArrowPath(rect, ArrowDirection.Up, ShaftAdj(adjustmentValues), HeadAdj(adjustmentValues)),
                ShapePresetType.DownArrow => ArrowPath(rect, ArrowDirection.Down, ShaftAdj(adjustmentValues), HeadAdj(adjustmentValues)),
                ShapePresetType.LeftRightArrow => DoubleArrowPath(rect, horizontal: true, ShaftAdj(adjustmentValues), DoubleHeadAdj(adjustmentValues)),
                ShapePresetType.UpDownArrow => DoubleArrowPath(rect, horizontal: false, ShaftAdj(adjustmentValues), DoubleHeadAdj(adjustmentValues)),
                ShapePresetType.WedgeRectCallout => WedgeCalloutPath(rect, BodyKind.Rect, adjustmentValues),
                ShapePresetType.WedgeRoundRectCallout => WedgeCalloutPath(rect, BodyKind.RoundRect, adjustmentValues),
                ShapePresetType.WedgeEllipseCallout => WedgeCalloutPath(rect, BodyKind.Ellipse, adjustmentValues),
                _ => RectPath(rect),
            };

        private static double ShaftAdj(IReadOnlyList<double> values) => Adj(values, 0, DefaultArrowShaftAdj);

        private static double HeadAdj(IReadOnlyList<double> values) => Adj(values, 1, DefaultArrowHeadAdj);

        private static double DoubleHeadAdj(IReadOnlyList<double> values) => Adj(values, 1, DefaultDoubleArrowHeadAdj);

        /// <summary>指定位置の調整値。ファイルに指定が無い(NaN)場合や範囲外は既定値を返す。</summary>
        private static double Adj(IReadOnlyList<double> values, int index, double defaultValue)
        {
            if (index >= values.Count || double.IsNaN(values[index]))
            {
                return defaultValue;
            }

            return values[index];
        }

        private static SKPath RectPath(SKRect rect)
        {
            var path = new SKPath();
            path.AddRect(rect);
            return path;
        }

        private static SKPath EllipsePath(SKRect rect)
        {
            var path = new SKPath();
            path.AddOval(rect);
            return path;
        }

        private static SKPath RoundRectPath(SKRect rect, double adj)
        {
            var radius = (float)(Math.Min(rect.Width, rect.Height) * Math.Max(0.0, Math.Min(0.5, adj)));
            var path = new SKPath();
            path.AddRoundRect(rect, radius, radius);
            return path;
        }

        /// <summary>頂角がX軸上を<paramref name="apexAdj"/>(0〜1、既定0.5=中央)の位置にある二等辺三角形。</summary>
        private static SKPath TrianglePath(SKRect rect, double apexAdj)
        {
            var apexX = rect.Left + (rect.Width * (float)Math.Max(0.0, Math.Min(1.0, apexAdj)));

            var path = new SKPath();
            path.MoveTo(apexX, rect.Top);
            path.LineTo(rect.Right, rect.Bottom);
            path.LineTo(rect.Left, rect.Bottom);
            path.Close();
            return path;
        }

        private enum ArrowDirection
        {
            Right,
            Left,
            Up,
            Down,
        }

        /// <summary>
        /// 単方向矢印。<paramref name="shaftAdj"/>は軸の太さ比(進行方向と垂直な辺に対する比率)、
        /// <paramref name="headAdj"/>は矢尻の長さ比(進行方向の辺に対する比率)。
        /// </summary>
        private static SKPath ArrowPath(SKRect rect, ArrowDirection direction, double shaftAdj, double headAdj)
        {
            // 右向きを基準に、幅=進行方向、高さ=垂直方向として組み立ててから座標を差し替える。
            var w = direction is ArrowDirection.Up or ArrowDirection.Down ? rect.Height : rect.Width;
            var h = direction is ArrowDirection.Up or ArrowDirection.Down ? rect.Width : rect.Height;

            var shaftHalf = (float)(h * Math.Max(0.0, Math.Min(1.0, shaftAdj)) / 2.0);
            var headLength = (float)(w * Math.Max(0.0, Math.Min(1.0, headAdj)));

            var midY = h / 2f;
            var shaftTop = midY - shaftHalf;
            var shaftBottom = midY + shaftHalf;
            var headStartX = w - headLength;

            // ローカル座標系(0,0)-(w,h)。右向き矢印として組み立てる。
            var local = new SKPath();
            local.MoveTo(0f, shaftTop);
            local.LineTo(headStartX, shaftTop);
            local.LineTo(headStartX, 0f);
            local.LineTo(w, midY);
            local.LineTo(headStartX, h);
            local.LineTo(headStartX, shaftBottom);
            local.LineTo(0f, shaftBottom);
            local.Close();

            return TransformArrowLocalPath(local, rect, direction);
        }

        /// <summary>両端に矢尻を持つ双方向矢印。</summary>
        private static SKPath DoubleArrowPath(SKRect rect, bool horizontal, double shaftAdj, double headAdj)
        {
            var w = horizontal ? rect.Width : rect.Height;
            var h = horizontal ? rect.Height : rect.Width;

            var shaftHalf = (float)(h * Math.Max(0.0, Math.Min(1.0, shaftAdj)) / 2.0);
            var headLength = (float)(w * Math.Max(0.0, Math.Min(0.5, headAdj)));

            var midY = h / 2f;
            var shaftTop = midY - shaftHalf;
            var shaftBottom = midY + shaftHalf;

            var local = new SKPath();
            local.MoveTo(0f, midY);
            local.LineTo(headLength, 0f);
            local.LineTo(headLength, shaftTop);
            local.LineTo(w - headLength, shaftTop);
            local.LineTo(w - headLength, 0f);
            local.LineTo(w, midY);
            local.LineTo(w - headLength, h);
            local.LineTo(w - headLength, shaftBottom);
            local.LineTo(headLength, shaftBottom);
            local.LineTo(headLength, h);
            local.Close();

            return TransformArrowLocalPath(local, rect, horizontal ? ArrowDirection.Right : ArrowDirection.Down);
        }

        /// <summary>
        /// 「右向き・幅=進行方向」のローカル座標(0,0)-(w,h)で組み立てたパスを、実際の矢印の向きに
        /// 応じて <paramref name="rect"/> 上へ配置する。各成分は明示的なアフィン変換で導出する
        /// (Right: 平行移動のみ。Left: 進行方向を反転。Up/Down: 進行方向とページのX/Yを入れ替える)。
        /// </summary>
        private static SKPath TransformArrowLocalPath(SKPath local, SKRect rect, ArrowDirection direction)
        {
            // SKMatrix(scaleX, skewX, transX, skewY, scaleY, transY, persp0, persp1, persp2):
            // X' = scaleX*x + skewX*y + transX, Y' = skewY*x + scaleY*y + transY
            var matrix = direction switch
            {
                // 右向き: ローカル(x, y) → (rect.Left + x, rect.Top + y)。
                ArrowDirection.Right => new SKMatrix(1f, 0f, rect.Left, 0f, 1f, rect.Top, 0f, 0f, 1f),
                // 左向き: 進行方向(x)を反転してrect.Rightを起点にする。
                ArrowDirection.Left => new SKMatrix(-1f, 0f, rect.Right, 0f, 1f, rect.Top, 0f, 0f, 1f),
                // 上向き: ローカルのx(進行方向)→ページのYを反転して上へ、ローカルのy(垂直方向)→ページのX。
                ArrowDirection.Up => new SKMatrix(0f, 1f, rect.Left, -1f, 0f, rect.Bottom, 0f, 0f, 1f),
                // 下向き: ローカルのx(進行方向)→ページのYをそのまま下へ、ローカルのy(垂直方向)→ページのX。
                ArrowDirection.Down => new SKMatrix(0f, 1f, rect.Left, 1f, 0f, rect.Top, 0f, 0f, 1f),
                _ => new SKMatrix(1f, 0f, rect.Left, 0f, 1f, rect.Top, 0f, 0f, 1f),
            };

            local.Transform(matrix);
            return local;
        }

        private enum BodyKind
        {
            Rect,
            RoundRect,
            Ellipse,
        }

        /// <summary>
        /// 吹き出し(本体形状 + 引き出し三角形)。<paramref name="adjustmentValues"/>の2値を
        /// 引き出し先端の位置(本体の幅・高さに対する比率。0〜1の外側も取りうる)として扱う。
        /// </summary>
        private static SKPath WedgeCalloutPath(SKRect rect, BodyKind body, IReadOnlyList<double> adjustmentValues)
        {
            var tipXAdj = Adj(adjustmentValues, 0, DefaultCalloutTipXAdj);
            var tipYAdj = Adj(adjustmentValues, 1, DefaultCalloutTipYAdj);

            var path = body switch
            {
                BodyKind.RoundRect => RoundRectPath(rect, DefaultRoundRectAdj),
                BodyKind.Ellipse => EllipsePath(rect),
                _ => RectPath(rect),
            };

            var tipX = rect.Left + (rect.Width * (float)tipXAdj);
            var tipY = rect.Top + (rect.Height * (float)tipYAdj);
            var baseSize = Math.Min(rect.Width, rect.Height) * 0.15f;

            SKPoint baseA, baseB;
            if (tipX < rect.Left)
            {
                baseA = new SKPoint(rect.Left, Clamp(tipY - baseSize, rect.Top, rect.Bottom));
                baseB = new SKPoint(rect.Left, Clamp(tipY + baseSize, rect.Top, rect.Bottom));
            }
            else if (tipX > rect.Right)
            {
                baseA = new SKPoint(rect.Right, Clamp(tipY - baseSize, rect.Top, rect.Bottom));
                baseB = new SKPoint(rect.Right, Clamp(tipY + baseSize, rect.Top, rect.Bottom));
            }
            else if (tipY < rect.Top)
            {
                baseA = new SKPoint(Clamp(tipX - baseSize, rect.Left, rect.Right), rect.Top);
                baseB = new SKPoint(Clamp(tipX + baseSize, rect.Left, rect.Right), rect.Top);
            }
            else
            {
                baseA = new SKPoint(Clamp(tipX - baseSize, rect.Left, rect.Right), rect.Bottom);
                baseB = new SKPoint(Clamp(tipX + baseSize, rect.Left, rect.Right), rect.Bottom);
            }

            var tail = new SKPath();
            tail.MoveTo(baseA);
            tail.LineTo(tipX, tipY);
            tail.LineTo(baseB);
            tail.Close();

            path.AddPath(tail);
            path.FillType = SKPathFillType.Winding;
            return path;
        }

        private static float Clamp(float value, float min, float max) => Math.Max(min, Math.Min(max, value));
    }
}
