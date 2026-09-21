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

        /// <summary>
        /// 吹き出しの引き出し先端の調整値(本体の幅・高さに対する比率)の絶対値の上限。
        /// ファイル由来の値をそのまま使うと極端な座標になりうるため、本体の数倍程度まで
        /// 外側にはみ出す吹き出しは許容しつつ有限の妥当な範囲に収める(security-reviewer指摘)。
        /// </summary>
        private const double CalloutTipAdjLimit = 5.0;

        /// <summary>
        /// star4/5/6/8共通の内側頂点の半径比(外接円半径に対する比率)の既定値。
        /// ECMA-376の一次資料への当たり直しはできていない暫定値(design.md参照)。
        /// </summary>
        private const double DefaultStarInnerRadiusRatio = 0.38;

        /// <summary>flowChartInputOutput(平行四辺形)の上下辺のずらし幅(矩形の幅に対する比率)。</summary>
        private const double InputOutputSkewRatio = 0.2;

        /// <summary>flowChartDocumentの波形の深さ(矩形の高さに対する比率)。</summary>
        private const double DocumentWaveDepthRatio = 0.08;

        /// <summary>flowChartPredefinedProcessの左右の縦線の位置(矩形の幅に対する内側からの比率)。</summary>
        private const double PredefinedProcessInsetRatio = 0.1;

        /// <summary>cloudCalloutを近似する円(バンプ)の個数。調整ガイドには対応しない固定値。</summary>
        private const int CloudBumpCount = 12;

        /// <summary>cloudCalloutのバンプ1つの半径(矩形の短辺に対する比率)。</summary>
        private const double CloudBumpRadiusRatio = 0.16;

        /// <summary>callout1/2/3の引き出し線の始点(本体の幅に対する左端からの比率)。</summary>
        private const double LeaderStartXRatio = 0.25;

        /// <summary>callout1/2/3の引き出し線の先端(本体の幅・高さに対する比率。0〜1の外側)。</summary>
        private const double LeaderTipXRatio = -0.25;

        private const double LeaderTipYRatio = 1.75;

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
                ShapePresetType.CloudCallout => CloudCalloutPath(rect),
                ShapePresetType.Callout1 => RectPath(rect),
                ShapePresetType.Callout2 => RectPath(rect),
                ShapePresetType.Callout3 => RectPath(rect),
                ShapePresetType.Star4 => StarPath(rect, 4, StarInnerRadiusRatio(adjustmentValues)),
                ShapePresetType.Star5 => StarPath(rect, 5, StarInnerRadiusRatio(adjustmentValues)),
                ShapePresetType.Star6 => StarPath(rect, 6, StarInnerRadiusRatio(adjustmentValues)),
                ShapePresetType.Star8 => StarPath(rect, 8, StarInnerRadiusRatio(adjustmentValues)),
                ShapePresetType.FlowChartProcess => RectPath(rect),
                ShapePresetType.FlowChartDecision => DiamondPath(rect),
                ShapePresetType.FlowChartTerminator => StadiumPath(rect),
                ShapePresetType.FlowChartInputOutput => ParallelogramPath(rect, InputOutputSkewRatio),
                ShapePresetType.FlowChartDocument => DocumentPath(rect),
                ShapePresetType.FlowChartPredefinedProcess => PredefinedProcessPath(rect),
                ShapePresetType.FlowChartConnector => EllipsePath(rect),
                _ => RectPath(rect),
            };

        /// <summary>
        /// 塗りつぶし用のジオメトリとは別に、枠線用のジオメトリを返す。callout1/2/3は本体(矩形)に
        /// 加えて塗りつぶしを持たない引き出し線(開いた折れ線)を枠線側にのみ追加するため、
        /// <see cref="Build"/>(本体のみ)とは異なるパスが必要になる。それ以外のプリセットは
        /// <see cref="Build"/>と同じジオメトリを枠線にも使う。
        /// </summary>
        public static SKPath BuildOutline(ShapePresetType preset, IReadOnlyList<double> adjustmentValues, SKRect rect) =>
            preset switch
            {
                ShapePresetType.Callout1 => CalloutOutlinePath(rect, 1),
                ShapePresetType.Callout2 => CalloutOutlinePath(rect, 2),
                ShapePresetType.Callout3 => CalloutOutlinePath(rect, 3),
                _ => Build(preset, adjustmentValues, rect),
            };

        private static double StarInnerRadiusRatio(IReadOnlyList<double> values) => Adj(values, 0, DefaultStarInnerRadiusRatio);

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
            // 他の調整値と異なり、引き出し先端は0〜1の範囲外(本体の外)を指すのが通常のため
            // 上限で丸めはしないが、ファイル由来の値(理論上はInt32の全域を100000で割った
            // 値域まで取りうる)をそのまま使うと極端に大きな座標がSKPathに渡ってしまう
            // (security-reviewer指摘)。本体の外側へ大きくはみ出す吹き出しを許容しつつ
            // 座標を有限の妥当な範囲に収めるため、絶対値で適度な上限にクランプする。
            var tipXAdj = Math.Max(-CalloutTipAdjLimit, Math.Min(CalloutTipAdjLimit, Adj(adjustmentValues, 0, DefaultCalloutTipXAdj)));
            var tipYAdj = Math.Max(-CalloutTipAdjLimit, Math.Min(CalloutTipAdjLimit, Adj(adjustmentValues, 1, DefaultCalloutTipYAdj)));

            var path = body switch
            {
                BodyKind.RoundRect => RoundRectPath(rect, DefaultRoundRectAdj),
                BodyKind.Ellipse => EllipsePath(rect),
                _ => RectPath(rect),
            };

            AddWedgeTail(path, rect, tipXAdj, tipYAdj);
            return path;
        }

        /// <summary>
        /// <paramref name="path"/>に、本体の輪郭外(または内)の1点(<paramref name="tipXAdj"/>/
        /// <paramref name="tipYAdj"/>で指定)へ向けた引き出し三角形を追加する
        /// (wedge系の吹き出し・cloudCalloutで共通)。
        /// </summary>
        private static void AddWedgeTail(SKPath path, SKRect rect, double tipXAdj, double tipYAdj)
        {
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

            using var tail = new SKPath();
            tail.MoveTo(baseA);
            tail.LineTo(tipX, tipY);
            tail.LineTo(baseB);
            tail.Close();

            path.AddPath(tail);
            path.FillType = SKPathFillType.Winding;
        }

        /// <summary>
        /// 雲形吹き出し(cloudCallout)。楕円本体の輪郭に沿って並べた円(バンプ)の和集合で
        /// 近似したシルエットに、wedgeEllipseCalloutと同じ引き出し三角形を追加する
        /// (design.md参照。個数・半径は固定値で調整ガイドには対応しない)。
        /// </summary>
        private static SKPath CloudCalloutPath(SKRect rect)
        {
            var centerX = (rect.Left + rect.Right) / 2f;
            var centerY = (rect.Top + rect.Bottom) / 2f;
            var radiusX = rect.Width / 2f;
            var radiusY = rect.Height / 2f;
            var bumpRadius = (float)(Math.Min(rect.Width, rect.Height) * CloudBumpRadiusRatio);

            SKPath? cloud = null;
            for (var i = 0; i < CloudBumpCount; i++)
            {
                var angle = 2.0 * Math.PI * i / CloudBumpCount;
                var bumpX = centerX + (radiusX * (float)Math.Cos(angle));
                var bumpY = centerY + (radiusY * (float)Math.Sin(angle));

                using var bump = new SKPath();
                bump.AddOval(new SKRect(bumpX - bumpRadius, bumpY - bumpRadius, bumpX + bumpRadius, bumpY + bumpRadius));

                if (cloud is null)
                {
                    cloud = new SKPath(bump);
                    continue;
                }

                using var previous = cloud;
                cloud = previous.Op(bump, SKPathOp.Union) ?? previous;
            }

            var result = cloud ?? EllipsePath(rect);
            AddWedgeTail(result, rect, DefaultCalloutTipXAdj, DefaultCalloutTipYAdj);
            return result;
        }

        /// <summary>星形(star4/5/6/8)。外接円の半径と内側頂点の半径比から交互に結んだ2N角形。</summary>
        private static SKPath StarPath(SKRect rect, int points, double innerRadiusRatio)
        {
            var centerX = (rect.Left + rect.Right) / 2f;
            var centerY = (rect.Top + rect.Bottom) / 2f;
            var outerRadius = Math.Min(rect.Width, rect.Height) / 2f;
            var innerRadius = outerRadius * (float)Math.Max(0.0, Math.Min(1.0, innerRadiusRatio));

            var path = new SKPath();
            var vertexCount = points * 2;
            for (var i = 0; i < vertexCount; i++)
            {
                // 最初の外側頂点を真上(-90度)に置き、外側・内側の頂点を交互に配置する。
                var angle = (-Math.PI / 2.0) + (i * Math.PI / points);
                var radius = i % 2 == 0 ? outerRadius : innerRadius;
                var x = centerX + (float)(radius * Math.Cos(angle));
                var y = centerY + (float)(radius * Math.Sin(angle));

                if (i == 0)
                {
                    path.MoveTo(x, y);
                }
                else
                {
                    path.LineTo(x, y);
                }
            }

            path.Close();
            return path;
        }

        /// <summary>flowChartDecision(判断)。矩形の上下左右の中点を結んだ菱形。</summary>
        private static SKPath DiamondPath(SKRect rect)
        {
            var midX = (rect.Left + rect.Right) / 2f;
            var midY = (rect.Top + rect.Bottom) / 2f;

            var path = new SKPath();
            path.MoveTo(midX, rect.Top);
            path.LineTo(rect.Right, midY);
            path.LineTo(midX, rect.Bottom);
            path.LineTo(rect.Left, midY);
            path.Close();
            return path;
        }

        /// <summary>flowChartTerminator(端子)。左右端を半円にした「スタジアム」形状。</summary>
        private static SKPath StadiumPath(SKRect rect)
        {
            var radius = Math.Min(rect.Width, rect.Height) / 2f;
            var path = new SKPath();
            path.AddRoundRect(rect, radius, radius);
            return path;
        }

        /// <summary>flowChartInputOutput(入出力)。上下の辺を左右にずらした平行四辺形。</summary>
        private static SKPath ParallelogramPath(SKRect rect, double skewRatio)
        {
            var skew = (float)(rect.Width * Math.Max(0.0, Math.Min(0.5, skewRatio)));

            var path = new SKPath();
            path.MoveTo(rect.Left + skew, rect.Top);
            path.LineTo(rect.Right, rect.Top);
            path.LineTo(rect.Right - skew, rect.Bottom);
            path.LineTo(rect.Left, rect.Bottom);
            path.Close();
            return path;
        }

        /// <summary>flowChartDocument(書類)。矩形の下辺を1つの緩やかな凹みにした形状。</summary>
        private static SKPath DocumentPath(SKRect rect)
        {
            var waveDepth = (float)(rect.Height * DocumentWaveDepthRatio);
            var midX = (rect.Left + rect.Right) / 2f;
            var baseY = rect.Bottom - waveDepth;

            var path = new SKPath();
            path.MoveTo(rect.Left, rect.Top);
            path.LineTo(rect.Right, rect.Top);
            path.LineTo(rect.Right, baseY);
            path.QuadTo(new SKPoint(midX, baseY + (waveDepth * 2f)), new SKPoint(rect.Left, baseY));
            path.Close();
            return path;
        }

        /// <summary>
        /// flowChartPredefinedProcess(定義済み処理)。矩形に加え、左右の辺の内側に縦線を1本ずつ追加する。
        /// 追加する2本の縦線は2頂点のみの退化した(面積0の)サブパスのため、塗りつぶし時には
        /// 何も描画されず、枠線として描画したときのみ見える。
        /// </summary>
        private static SKPath PredefinedProcessPath(SKRect rect)
        {
            var inset = (float)(rect.Width * PredefinedProcessInsetRatio);

            var path = RectPath(rect);
            path.MoveTo(rect.Left + inset, rect.Top);
            path.LineTo(rect.Left + inset, rect.Bottom);
            path.MoveTo(rect.Right - inset, rect.Top);
            path.LineTo(rect.Right - inset, rect.Bottom);
            return path;
        }

        /// <summary>callout1/2/3の枠線用ジオメトリ。本体(矩形)に、塗りつぶしを持たないN本の引き出し折れ線を加える。</summary>
        private static SKPath CalloutOutlinePath(SKRect rect, int segments)
        {
            var path = RectPath(rect);
            var points = BuildLeaderPoints(rect, segments);

            path.MoveTo(points[0]);
            for (var i = 1; i < points.Length; i++)
            {
                path.LineTo(points[i]);
            }

            return path;
        }

        /// <summary>
        /// callout1/2/3の引き出し線の頂点列(本体上の始点 → <paramref name="segments"/> - 1個の
        /// 折れ点 → 本体外側の先端)。既定で左下方向へ引き出す(wedge系の既定方向と同じ慣習)。
        /// </summary>
        private static SKPoint[] BuildLeaderPoints(SKRect rect, int segments)
        {
            var startX = rect.Left + (rect.Width * (float)LeaderStartXRatio);
            var startY = rect.Bottom;
            var tipX = rect.Left + (rect.Width * (float)LeaderTipXRatio);
            var tipY = rect.Top + (rect.Height * (float)LeaderTipYRatio);

            var points = new SKPoint[segments + 1];
            points[0] = new SKPoint(startX, startY);
            points[segments] = new SKPoint(tipX, tipY);

            for (var i = 1; i < segments; i++)
            {
                var t = (float)i / segments;
                var bendX = startX + ((tipX - startX) * t);
                var bendY = startY + ((tipY - startY) * t);

                // 直線的な等分点のままだと折れ線に見えないため、偶奇でわずかにずらす。
                bendY += (i % 2 == 0 ? -1f : 1f) * (float)(rect.Height * 0.05);
                points[i] = new SKPoint(bendX, bendY);
            }

            return points;
        }

        private static float Clamp(float value, float min, float max) => Math.Max(min, Math.Min(max, value));
    }
}
