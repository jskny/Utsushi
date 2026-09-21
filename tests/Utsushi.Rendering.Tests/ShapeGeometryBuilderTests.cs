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
