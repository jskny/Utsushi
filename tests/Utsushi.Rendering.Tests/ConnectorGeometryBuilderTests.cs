using SkiaSharp;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// 対応済み接続線プリセット(要件10.9補足)のパス生成の検証。
    /// </summary>
    public sealed class ConnectorGeometryBuilderTests
    {
        private static readonly SKRect Rect = new(10, 20, 110, 70); // 幅100 x 高さ50

        [Theory]
        [InlineData(ConnectorPresetType.Straight)]
        [InlineData(ConnectorPresetType.Bent2Segment)]
        [InlineData(ConnectorPresetType.Bent3Segment)]
        [InlineData(ConnectorPresetType.Curved2Segment)]
        [InlineData(ConnectorPresetType.Curved3Segment)]
        public void 全プリセットで空でない開いたパスを生成する(ConnectorPresetType preset)
        {
            using var path = ConnectorGeometryBuilder.Build(preset, flipHorizontal: false, flipVertical: false, Rect);

            Assert.False(path.IsEmpty);

            // 接続線は塗りつぶし無しの開いたパス(design.md参照)。SKPathでは、明示的にCloseを
            // 呼んでいない限りFillTypeにかかわらずパスは開いたまま(始点と終点が異なる)であるため、
            // 最初と最後の点が一致しないことで「閉じていない」ことを確認する。
            Assert.NotEqual(path.Points[0], path.Points[^1]);
        }

        [Fact]
        public void Straightは反転が無い場合は左上から右下への直線になる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Straight, flipHorizontal: false, flipVertical: false, Rect);

            Assert.Equal(2, path.PointCount);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Top), path.Points[0]);
            Assert.Equal(new SKPoint(Rect.Right, Rect.Bottom), path.Points[1]);
        }

        [Fact]
        public void Straightは左右反転すると始点終点のX座標が入れ替わる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Straight, flipHorizontal: true, flipVertical: false, Rect);

            Assert.Equal(new SKPoint(Rect.Right, Rect.Top), path.Points[0]);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Bottom), path.Points[1]);
        }

        [Fact]
        public void Straightは上下反転すると始点終点のY座標が入れ替わる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Straight, flipHorizontal: false, flipVertical: true, Rect);

            Assert.Equal(new SKPoint(Rect.Left, Rect.Bottom), path.Points[0]);
            Assert.Equal(new SKPoint(Rect.Right, Rect.Top), path.Points[1]);
        }

        [Fact]
        public void Straightは上下左右反転すると右下から左上への直線になる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Straight, flipHorizontal: true, flipVertical: true, Rect);

            Assert.Equal(new SKPoint(Rect.Right, Rect.Bottom), path.Points[0]);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Top), path.Points[1]);
        }

        [Fact]
        public void Bent2Segmentは水平垂直の順で1回折れる2辺になる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Bent2Segment, flipHorizontal: false, flipVertical: false, Rect);

            Assert.Equal(3, path.PointCount);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Top), path.Points[0]);
            // 折れ点は終点のXと始点のYを持つ(水平→垂直の順)。
            Assert.Equal(new SKPoint(Rect.Right, Rect.Top), path.Points[1]);
            Assert.Equal(new SKPoint(Rect.Right, Rect.Bottom), path.Points[2]);
        }

        [Fact]
        public void Bent2Segmentは反転すると折れ点の位置も反転後の始点終点に追従する()
        {
            using var flipped = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Bent2Segment, flipHorizontal: true, flipVertical: false, Rect);

            Assert.Equal(new SKPoint(Rect.Right, Rect.Top), flipped.Points[0]);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Top), flipped.Points[1]);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Bottom), flipped.Points[2]);
        }

        [Fact]
        public void Bent3Segmentは始点終点のX中央で2回折れる3辺になる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Bent3Segment, flipHorizontal: false, flipVertical: false, Rect);

            Assert.Equal(4, path.PointCount);
            var midX = (Rect.Left + Rect.Right) / 2f;
            Assert.Equal(new SKPoint(Rect.Left, Rect.Top), path.Points[0]);
            Assert.Equal(new SKPoint(midX, Rect.Top), path.Points[1]);
            Assert.Equal(new SKPoint(midX, Rect.Bottom), path.Points[2]);
            Assert.Equal(new SKPoint(Rect.Right, Rect.Bottom), path.Points[3]);
        }

        [Fact]
        public void Curved2Segmentは始点終点を結ぶ1本の2次ベジェ曲線になる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Curved2Segment, flipHorizontal: false, flipVertical: false, Rect);

            // MoveTo + QuadTo(制御点+終点)で3点。
            Assert.Equal(3, path.PointCount);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Top), path.Points[0]);
            Assert.Equal(new SKPoint(Rect.Right, Rect.Bottom), path.Points[^1]);

            // 開いた曲線であり、始点・終点はRectの対角(反転無しなので左上・右下)。
            Assert.True(path.Bounds.Left <= Rect.Left + 0.01f);
            Assert.True(path.Bounds.Top <= Rect.Top + 0.01f);
        }

        [Fact]
        public void Curved3SegmentはBent3Segmentと同じ折れ点を通る2本のベジェ曲線になる()
        {
            using var path = ConnectorGeometryBuilder.Build(
                ConnectorPresetType.Curved3Segment, flipHorizontal: false, flipVertical: false, Rect);

            // MoveTo + QuadTo*2で5点(各QuadToは制御点+終点の2点を追加)。
            Assert.Equal(5, path.PointCount);
            Assert.Equal(new SKPoint(Rect.Left, Rect.Top), path.Points[0]);
            Assert.Equal(new SKPoint(Rect.Right, Rect.Bottom), path.Points[^1]);
        }
    }
}
