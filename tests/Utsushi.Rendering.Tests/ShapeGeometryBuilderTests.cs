using System;
using System.Collections.Generic;
using System.Linq;
using SkiaSharp;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// 対応済みプリセットジオメトリ(要件10.1)のパス生成の検証。
    /// </summary>
    public sealed class ShapeGeometryBuilderTests
    {
        private static readonly SKRect Rect = new(10, 20, 110, 70); // 幅100 x 高さ50

        public static IEnumerable<object[]> AllPresets()
        {
            foreach (ShapePresetType preset in Enum.GetValues(typeof(ShapePresetType)))
            {
                yield return new object[] { preset };
            }
        }

        [Theory]
        [MemberData(nameof(AllPresets))]
        public void 全プリセットで空でないパスを生成する(ShapePresetType preset)
        {
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            Assert.False(path.IsEmpty);
        }

        [Theory]
        [InlineData(ShapePresetType.Rect)]
        [InlineData(ShapePresetType.Ellipse)]
        [InlineData(ShapePresetType.RoundRect)]
        [InlineData(ShapePresetType.Triangle)]
        [InlineData(ShapePresetType.RightArrow)]
        [InlineData(ShapePresetType.LeftArrow)]
        [InlineData(ShapePresetType.UpArrow)]
        [InlineData(ShapePresetType.DownArrow)]
        [InlineData(ShapePresetType.LeftRightArrow)]
        [InlineData(ShapePresetType.UpDownArrow)]
        public void 吹き出し以外は指定矩形の内側に収まる(ShapePresetType preset)
        {
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            AssertBoundsApproximately(Rect, path.Bounds);
        }

        [Theory]
        [InlineData(ShapePresetType.WedgeRectCallout)]
        [InlineData(ShapePresetType.WedgeRoundRectCallout)]
        [InlineData(ShapePresetType.WedgeEllipseCallout)]
        public void 吹き出しは引き出し先端ぶん本体矩形の外側へ広がる(ShapePresetType preset)
        {
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            // 既定の引き出し先端(adj1=-0.25)は本体の左側を指すため、左方向へ広がる。
            Assert.True(path.Bounds.Left < Rect.Left);
        }

        [Fact]
        public void 吹き出しの引き出し先端は極端な調整値でも有限の範囲にクランプされる()
        {
            // security-reviewer指摘の回帰テスト: a:gd/@fmlaは理論上Int32の全域を100000で
            // 割った値(最大約±21474.8)まで取りうるが、そのまま使うと座標が極端に大きくなる。
            using var path = ShapeGeometryBuilder.Build(
                ShapePresetType.WedgeRectCallout, new[] { 100000.0, -100000.0 }, Rect);

            Assert.False(float.IsNaN(path.Bounds.Left));
            Assert.False(float.IsInfinity(path.Bounds.Left));
            Assert.True(path.Bounds.Width < Rect.Width * 20, $"Width={path.Bounds.Width}");
            Assert.True(path.Bounds.Height < Rect.Height * 20, $"Height={path.Bounds.Height}");
        }

        [Fact]
        public void RightArrowの先端は矩形の右端にある()
        {
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.RightArrow, Array.Empty<double>(), Rect);

            Assert.Equal(Rect.Right, path.Bounds.Right, 3);
        }

        [Fact]
        public void LeftArrowの先端は矩形の左端にある()
        {
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.LeftArrow, Array.Empty<double>(), Rect);

            Assert.Equal(Rect.Left, path.Bounds.Left, 3);
        }

        [Fact]
        public void UpArrowの先端は矩形の上端にある()
        {
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.UpArrow, Array.Empty<double>(), Rect);

            Assert.Equal(Rect.Top, path.Bounds.Top, 3);
        }

        [Fact]
        public void DownArrowの先端は矩形の下端にある()
        {
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.DownArrow, Array.Empty<double>(), Rect);

            Assert.Equal(Rect.Bottom, path.Bounds.Bottom, 3);
        }

        [Fact]
        public void LeftRightArrowは両端に矢尻を持ち矩形いっぱいに広がる()
        {
            // 既定の矢尻長さ比(0.25)を使うと、両端の矢尻の間にシャフト(軸)が残る
            // (単方向矢印と同じ既定値0.5を使うと軸が消えて菱形に潰れる不具合があった。design.md参照)。
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.LeftRightArrow, Array.Empty<double>(), Rect);

            AssertBoundsApproximately(Rect, path.Bounds);
            Assert.True(path.PointCount >= 8, "両端に矢尻を持つ双方向矢印は8点以上の多角形になるはず");
        }

        [Theory]
        [InlineData(0.0)]
        [InlineData(0.5)]
        [InlineData(1.0)]
        public void Triangleの頂角は調整値に応じてX軸上を移動する(double apexAdj)
        {
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.Triangle, new[] { apexAdj }, Rect);

            // TrianglePathはMoveTo(apexX, rect.Top)から開始するため、先頭の点が頂角。
            var apex = path.Points[0];
            var expectedX = Rect.Left + (Rect.Width * (float)apexAdj);
            Assert.Equal(expectedX, apex.X, 3);
            Assert.Equal(Rect.Top, apex.Y, 3);
        }

        [Fact]
        public void 調整値がNaNの場合は既定値の頂角中央を使う()
        {
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.Triangle, new[] { double.NaN }, Rect);

            var apex = path.Points[0];
            Assert.Equal(Rect.Left + (Rect.Width / 2f), apex.X, 3);
        }

        [Fact]
        public void 調整値が空配列の場合も既定値を使う()
        {
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.Triangle, Array.Empty<double>(), Rect);

            var apex = path.Points[0];
            Assert.Equal(Rect.Left + (Rect.Width / 2f), apex.X, 3);
        }

        // -- 星形・フローチャート記号・吹き出し(新規プリセット) -------------------------------

        [Theory]
        [InlineData(ShapePresetType.Star4)]
        [InlineData(ShapePresetType.Star5)]
        [InlineData(ShapePresetType.Star6)]
        [InlineData(ShapePresetType.Star8)]
        public void 星形は指定矩形の内側に収まる(ShapePresetType preset)
        {
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            AssertBoundsWithin(Rect, path.Bounds);
        }

        [Theory]
        [InlineData(ShapePresetType.FlowChartProcess)]
        [InlineData(ShapePresetType.FlowChartDecision)]
        [InlineData(ShapePresetType.FlowChartTerminator)]
        [InlineData(ShapePresetType.FlowChartInputOutput)]
        [InlineData(ShapePresetType.FlowChartPredefinedProcess)]
        [InlineData(ShapePresetType.FlowChartConnector)]
        public void ほとんどのフローチャート記号は指定矩形とちょうど一致する境界になる(ShapePresetType preset)
        {
            // 処理/判断/端子/入出力/定義済み処理/結合子は、いずれも矩形の4辺(または中点)に
            // 頂点が接するため、パスの境界は入力矩形と完全に一致するはず。
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            AssertBoundsApproximately(Rect, path.Bounds);
        }

        [Fact]
        public void flowChartDocumentは下辺の波形ぶん矩形よりわずかに下へふくらむ()
        {
            // 書類の下辺は「緩やかな凹み」をQuadToの制御点で表現しており、制御点自体は
            // 矩形の下端よりさらに下(高さの8%ぶん)に置かれるため、パスの境界は
            // 矩形よりわずかに(意図的に)下へはみ出す。
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.FlowChartDocument, Array.Empty<double>(), Rect);

            Assert.Equal(Rect.Left, path.Bounds.Left, 3);
            Assert.Equal(Rect.Top, path.Bounds.Top, 3);
            Assert.Equal(Rect.Right, path.Bounds.Right, 3);
            Assert.True(path.Bounds.Bottom > Rect.Bottom, $"Bottom: expected > {Rect.Bottom} actual={path.Bounds.Bottom}");
            Assert.True(path.Bounds.Bottom < Rect.Bottom + Rect.Height, "はみ出しは矩形の高さを超えないはず(意図しない破綻の検出)");
        }

        [Theory]
        [InlineData(ShapePresetType.Star4, 0.25)]
        [InlineData(ShapePresetType.Star5, 0.382)]
        [InlineData(ShapePresetType.Star6, 0.577)]
        [InlineData(ShapePresetType.Star8, 0.75)]
        public void 星形の既定内側半径比はプリセットごとの規定値になる(ShapePresetType preset, double expectedRatio)
        {
            // 要件10.12: 内側頂点の半径比の既定値はECMA-376のadj既定値(0〜50000)を
            // 50000で割った比率であり、star4/5/6/8で共通ではない。
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            var centerX = (Rect.Left + Rect.Right) / 2f;
            var centerY = (Rect.Top + Rect.Bottom) / 2f;
            var outerRadius = Math.Min(Rect.Width, Rect.Height) / 2f;

            // StarPathはi=0(外側)から始まり、奇数インデックス(i=1)が最初の内側頂点になる。
            var innerVertex = path.Points[1];
            var actualInnerRadius = Distance(innerVertex, centerX, centerY);
            var expectedInnerRadius = outerRadius * (float)expectedRatio;

            Assert.Equal(expectedInnerRadius, actualInnerRadius, 2);
        }

        [Fact]
        public void 星形の既定内側半径比はstar4_5_6_8ですべて異なる()
        {
            // 単一の共通既定値(旧実装の0.38)へ後退していないことの回帰確認(要件10.12)。
            var presets = new[] { ShapePresetType.Star4, ShapePresetType.Star5, ShapePresetType.Star6, ShapePresetType.Star8 };
            var centerX = (Rect.Left + Rect.Right) / 2f;
            var centerY = (Rect.Top + Rect.Bottom) / 2f;

            var innerRadii = presets
                .Select(preset =>
                {
                    using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);
                    return Distance(path.Points[1], centerX, centerY);
                })
                .ToList();

            Assert.Equal(presets.Length, innerRadii.Distinct().Count());
        }

        [Fact]
        public void cloudCalloutの引き出し位置は調整値を明示指定すると変わる()
        {
            // 要件10.13: adj1(X方向)/adj2(Y方向)を明示指定すると引き出し三角形の先端位置が変わる。
            // 既定(adj1=-0.25)は本体の左外側を指すが、adj1=1.25を指定すると右外側を指すようになる。
            using var defaultPath = ShapeGeometryBuilder.Build(ShapePresetType.CloudCallout, Array.Empty<double>(), Rect);
            using var rightPath = ShapeGeometryBuilder.Build(ShapePresetType.CloudCallout, new[] { 1.25, 0.5 }, Rect);

            Assert.True(
                rightPath.Bounds.Right > defaultPath.Bounds.Right,
                $"Right: default={defaultPath.Bounds.Right} right={rightPath.Bounds.Right}");
            Assert.True(
                rightPath.Bounds.Left > defaultPath.Bounds.Left,
                $"Left: default={defaultPath.Bounds.Left} right={rightPath.Bounds.Left}");
        }

        private static float Distance(SKPoint point, float centerX, float centerY) =>
            (float)Math.Sqrt(Math.Pow(point.X - centerX, 2) + Math.Pow(point.Y - centerY, 2));

        [Fact]
        public void cloudCalloutは楕円本体に加え引き出し先端ぶん本体矩形の外側へ広がる()
        {
            // wedge系の吹き出しと同じ既定の引き出し先端(adj1=-0.25)を使うため、左方向へ広がる
            // (design.md参照。バンプの個数・半径は固定値で調整ガイドには対応しない)。
            using var path = ShapeGeometryBuilder.Build(ShapePresetType.CloudCallout, Array.Empty<double>(), Rect);

            Assert.False(path.IsEmpty);
            Assert.True(path.Bounds.Left < Rect.Left, $"Left: expected < {Rect.Left} actual={path.Bounds.Left}");
            // 引き出し先端以外(雲本体)は概ね矩形内に収まる想定であり、無制限にはみ出さない。
            Assert.True(path.Bounds.Width < Rect.Width * 3, $"Width={path.Bounds.Width}");
            Assert.True(path.Bounds.Height < Rect.Height * 3, $"Height={path.Bounds.Height}");
        }

        public static IEnumerable<object[]> LineCallouts()
        {
            // (プリセット, 折れ数, 本体枠線の有無, 強調線の有無)
            yield return new object[] { ShapePresetType.Callout1, 1, false, false };
            yield return new object[] { ShapePresetType.Callout2, 2, false, false };
            yield return new object[] { ShapePresetType.Callout3, 3, false, false };
            yield return new object[] { ShapePresetType.BorderCallout1, 1, true, false };
            yield return new object[] { ShapePresetType.BorderCallout2, 2, true, false };
            yield return new object[] { ShapePresetType.BorderCallout3, 3, true, false };
            yield return new object[] { ShapePresetType.AccentCallout1, 1, false, true };
            yield return new object[] { ShapePresetType.AccentCallout2, 2, false, true };
            yield return new object[] { ShapePresetType.AccentCallout3, 3, false, true };
            yield return new object[] { ShapePresetType.AccentBorderCallout1, 1, true, true };
            yield return new object[] { ShapePresetType.AccentBorderCallout2, 2, true, true };
            yield return new object[] { ShapePresetType.AccentBorderCallout3, 3, true, true };
        }

        public static IEnumerable<object[]> LineCalloutPresets() => LineCallouts().Select(row => new[] { row[0] });

        [Theory]
        [MemberData(nameof(LineCalloutPresets))]
        public void 線吹き出しのBuildは本体矩形のみで引き出し線を含まない(ShapePresetType preset)
        {
            // Build(塗りつぶし用)は本体(矩形)のみであり、引き出し線(枠線専用)は含まない設計
            // (design.md「no-fillの引き出し線」要件)。境界は入力矩形とちょうど一致するはず。
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            AssertBoundsApproximately(Rect, path.Bounds);
        }

        [Theory]
        [MemberData(nameof(LineCallouts))]
        public void 線吹き出しのBuildOutlineは種類ごとに本体枠線と強調線と引き出し線を持つ(ShapePresetType preset, int segments, bool hasBorder, bool hasAccentBar)
        {
            using var outline = ShapeGeometryBuilder.BuildOutline(preset, Array.Empty<double>(), Rect);

            // 本体枠線は矩形(4点)、強調線は縦線(2点)、引き出し線は segments + 1 点の折れ線。
            var expectedPoints = (hasBorder ? 4 : 0) + (hasAccentBar ? 2 : 0) + segments + 1;
            Assert.Equal(expectedPoints, outline.PointCount);

            // ECMA-376既定の引き出し線は本体の左外側・下外側へ伸びる。
            Assert.True(outline.Bounds.Left < Rect.Left, $"Left: expected < {Rect.Left} actual={outline.Bounds.Left}");
            Assert.True(outline.Bounds.Bottom > Rect.Bottom, $"Bottom: expected > {Rect.Bottom} actual={outline.Bounds.Bottom}");
        }

        [Fact]
        public void 線吹き出しの引き出し線は調整値のyとxの組で頂点が決まる()
        {
            // borderCallout2: (adj1=y1, adj2=x1), (adj3=y2, adj4=x2), (adj5=y3, adj6=x3)。
            // 本体枠線(矩形4点)の後に引き出し線の3点が続く。
            var adj = new[] { 0.5, 1.0, 0.5, 1.5, 2.0, 2.0 };
            using var outline = ShapeGeometryBuilder.BuildOutline(ShapePresetType.BorderCallout2, adj, Rect);

            var points = outline.Points;
            Assert.Equal(7, points.Length);
            AssertPoint(Rect.Left + Rect.Width, Rect.Top + (Rect.Height * 0.5f), points[4]);
            AssertPoint(Rect.Left + (Rect.Width * 1.5f), Rect.Top + (Rect.Height * 0.5f), points[5]);
            AssertPoint(Rect.Left + (Rect.Width * 2f), Rect.Top + (Rect.Height * 2f), points[6]);
        }

        [Fact]
        public void 線吹き出しの調整値が無い位置は既定値で補う()
        {
            // adj1/adj2(始点)だけ指定し、残り(adj3/adj4)はファイルに無い(NaN)ケース。
            var adj = new[] { 0.0, 0.0, double.NaN, double.NaN };
            using var outline = ShapeGeometryBuilder.BuildOutline(ShapePresetType.Callout1, adj, Rect);

            var points = outline.Points;
            Assert.Equal(2, points.Length);
            AssertPoint(Rect.Left, Rect.Top, points[0]);

            // callout1の既定値: adj3=112500(y), adj4=-38333(x)
            AssertPoint(Rect.Left + (Rect.Width * -0.38333f), Rect.Top + (Rect.Height * 1.125f), points[1]);
        }

        [Fact]
        public void 強調線付き吹き出しの強調線は引き出し線の始点のX位置で本体の上端から下端まで引く()
        {
            var adj = new[] { 0.5, 0.25, 1.5, -0.5 };
            using var outline = ShapeGeometryBuilder.BuildOutline(ShapePresetType.AccentCallout1, adj, Rect);

            var points = outline.Points;
            var accentX = Rect.Left + (Rect.Width * 0.25f);
            AssertPoint(accentX, Rect.Top, points[0]);
            AssertPoint(accentX, Rect.Bottom, points[1]);
        }

        [Fact]
        public void 線吹き出しの極端な調整値は有限の範囲に収める()
        {
            var adj = new[] { 1e9, -1e9, 1e9, 1e9 };
            using var outline = ShapeGeometryBuilder.BuildOutline(ShapePresetType.BorderCallout1, adj, Rect);

            // 上限(本体の幅・高さの1000倍)で止まり、有限の座標に収まる。
            Assert.True(float.IsFinite(outline.Bounds.Width) && float.IsFinite(outline.Bounds.Height));
            Assert.True(outline.Bounds.Width <= Rect.Width * 2001, $"Width={outline.Bounds.Width}");
            Assert.True(outline.Bounds.Height <= Rect.Height * 2001, $"Height={outline.Bounds.Height}");
        }

        [Fact]
        public void 線吹き出しは小さな本体から離れたセルを指す実用的な調整値をそのまま使う()
        {
            // 本体の高さの7.5倍下・幅の10倍右を指す(小さなラベルから離れたセルを指す配置)。
            // wedge系の上限(±5倍)で丸めると先端が動いてしまう(layout-fidelity-reviewer指摘)。
            var adj = new[] { 0.5, 1.0, 7.5, 10.0 };
            using var outline = ShapeGeometryBuilder.BuildOutline(ShapePresetType.BorderCallout1, adj, Rect);

            var points = outline.Points;
            AssertPoint(Rect.Left + (Rect.Width * 10f), Rect.Top + (Rect.Height * 7.5f), points[5]);
        }

        public static IEnumerable<object[]> PresetsOtherThanCallouts()
        {
            var lineCallouts = new HashSet<ShapePresetType>();
            foreach (var row in LineCallouts())
            {
                lineCallouts.Add((ShapePresetType)row[0]);
            }

            foreach (ShapePresetType preset in Enum.GetValues(typeof(ShapePresetType)))
            {
                if (lineCallouts.Contains(preset))
                {
                    continue;
                }

                yield return new object[] { preset };
            }
        }

        [Theory]
        [MemberData(nameof(PresetsOtherThanCallouts))]
        public void 線吹き出し以外はBuildOutlineがBuildと同じジオメトリを返す(ShapePresetType preset)
        {
            using var body = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);
            using var outline = ShapeGeometryBuilder.BuildOutline(preset, Array.Empty<double>(), Rect);

            Assert.Equal(body.PointCount, outline.PointCount);
            AssertBoundsApproximately(body.Bounds, outline.Bounds);
        }

        private static void AssertPoint(float expectedX, float expectedY, SKPoint actual)
        {
            const float tolerance = 0.01f;
            Assert.True(Math.Abs(expectedX - actual.X) < tolerance, $"X: expected={expectedX} actual={actual.X}");
            Assert.True(Math.Abs(expectedY - actual.Y) < tolerance, $"Y: expected={expectedY} actual={actual.Y}");
        }

        private static void AssertBoundsWithin(SKRect outer, SKRect inner)
        {
            const float tolerance = 0.01f;
            Assert.True(inner.Left >= outer.Left - tolerance, $"Left: outer={outer.Left} inner={inner.Left}");
            Assert.True(inner.Top >= outer.Top - tolerance, $"Top: outer={outer.Top} inner={inner.Top}");
            Assert.True(inner.Right <= outer.Right + tolerance, $"Right: outer={outer.Right} inner={inner.Right}");
            Assert.True(inner.Bottom <= outer.Bottom + tolerance, $"Bottom: outer={outer.Bottom} inner={inner.Bottom}");
        }

        private static void AssertBoundsApproximately(SKRect expected, SKRect actual)
        {
            const float tolerance = 0.01f;
            Assert.True(Math.Abs(expected.Left - actual.Left) < tolerance, $"Left: expected={expected.Left} actual={actual.Left}");
            Assert.True(Math.Abs(expected.Top - actual.Top) < tolerance, $"Top: expected={expected.Top} actual={actual.Top}");
            Assert.True(Math.Abs(expected.Right - actual.Right) < tolerance, $"Right: expected={expected.Right} actual={actual.Right}");
            Assert.True(Math.Abs(expected.Bottom - actual.Bottom) < tolerance, $"Bottom: expected={expected.Bottom} actual={actual.Bottom}");
        }
    }
}
