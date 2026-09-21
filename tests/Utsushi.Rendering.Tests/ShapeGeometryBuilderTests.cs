using System;
using System.Collections.Generic;
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

        [Theory]
        [InlineData(ShapePresetType.Callout1)]
        [InlineData(ShapePresetType.Callout2)]
        [InlineData(ShapePresetType.Callout3)]
        public void callout系のBuildは本体矩形のみで引き出し線を含まない(ShapePresetType preset)
        {
            // Build(塗りつぶし用)は本体(矩形)のみであり、引き出し線(枠線専用)は含まない設計
            // (design.md「no-fillの引き出し線」要件)。境界は入力矩形とちょうど一致するはず。
            using var path = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);

            AssertBoundsApproximately(Rect, path.Bounds);
        }

        [Theory]
        [InlineData(ShapePresetType.Callout1, 1)]
        [InlineData(ShapePresetType.Callout2, 2)]
        [InlineData(ShapePresetType.Callout3, 3)]
        public void callout系のBuildOutlineはBuildと異なり引き出し線ぶん矩形の外側へ広がる(ShapePresetType preset, int segments)
        {
            using var body = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);
            using var outline = ShapeGeometryBuilder.BuildOutline(preset, Array.Empty<double>(), Rect);

            // Buildは本体(矩形)のみの4点(+Close)だが、BuildOutlineは本体に加えて
            // segments+1個の頂点からなる引き出し折れ線を追加で持つため点数が異なる。
            Assert.NotEqual(body.PointCount, outline.PointCount);
            Assert.True(outline.PointCount > body.PointCount + segments, "引き出し線の頂点が追加されているはず");

            // Buildは矩形の外に出ないが、BuildOutlineは引き出し先端(既定で左下方向)ぶん外側へ広がる。
            AssertBoundsApproximately(Rect, body.Bounds);
            Assert.True(outline.Bounds.Left < Rect.Left, $"Left: expected < {Rect.Left} actual={outline.Bounds.Left}");
            Assert.True(outline.Bounds.Bottom > Rect.Bottom, $"Bottom: expected > {Rect.Bottom} actual={outline.Bounds.Bottom}");
        }

        public static IEnumerable<object[]> PresetsOtherThanCallouts()
        {
            foreach (ShapePresetType preset in Enum.GetValues(typeof(ShapePresetType)))
            {
                if (preset is ShapePresetType.Callout1 or ShapePresetType.Callout2 or ShapePresetType.Callout3)
                {
                    continue;
                }

                yield return new object[] { preset };
            }
        }

        [Theory]
        [MemberData(nameof(PresetsOtherThanCallouts))]
        public void callout系以外はBuildOutlineがBuildと同じジオメトリを返す(ShapePresetType preset)
        {
            using var body = ShapeGeometryBuilder.Build(preset, Array.Empty<double>(), Rect);
            using var outline = ShapeGeometryBuilder.BuildOutline(preset, Array.Empty<double>(), Rect);

            Assert.Equal(body.PointCount, outline.PointCount);
            AssertBoundsApproximately(body.Bounds, outline.Bounds);
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
