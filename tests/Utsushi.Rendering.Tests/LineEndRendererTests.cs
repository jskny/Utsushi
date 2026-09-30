using System;
using System.Linq;
using SkiaSharp;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// 線の端の矢印(<c>a:headEnd</c>/<c>a:tailEnd</c>。要件10.18。design.md「矢印」)の描画の検証。
    /// </summary>
    /// <remarks>
    /// 200x200 のビットマップに黒で描き、不透明度が半分以上のピクセルの外接矩形と個々のピクセルで形と向きを確かめる。
    /// 矢印の大きさは線の太さの倍数(小=2、中=3、大=5)で、倍数を掛ける元の太さは最小1pt。
    /// </remarks>
    public sealed class LineEndRendererTests
    {
        private const int Size = 200;

        private static readonly SKPoint Tip = new(100, 100);

        private static readonly SKPoint Right = new(1, 0);

        private static LineEndStyle Medium(LineEndType type) => new(type, LineEndSize.Medium, LineEndSize.Medium);

        /// <summary><paramref name="draw"/> で描いたビットマップを返す。</summary>
        private static SKBitmap Render(Action<SKCanvas> draw)
        {
            var bitmap = new SKBitmap(new SKImageInfo(Size, Size, SKColorType.Rgba8888, SKAlphaType.Premul));
            using var canvas = new SKCanvas(bitmap);
            canvas.Clear(SKColors.Transparent);
            draw(canvas);
            canvas.Flush();
            return bitmap;
        }

        private static bool Inked(SKBitmap bitmap, int x, int y) => bitmap.GetPixel(x, y).Alpha >= 128;

        /// <summary>インクのあるピクセルの外接矩形(右端・下端はピクセルの右端・下端)。インクが無ければ null。</summary>
        private static SKRectI? InkBounds(SKBitmap bitmap)
        {
            int left = int.MaxValue, top = int.MaxValue, right = int.MinValue, bottom = int.MinValue;
            for (var y = 0; y < bitmap.Height; y++)
            {
                for (var x = 0; x < bitmap.Width; x++)
                {
                    if (!Inked(bitmap, x, y))
                    {
                        continue;
                    }

                    left = Math.Min(left, x);
                    top = Math.Min(top, y);
                    right = Math.Max(right, x + 1);
                    bottom = Math.Max(bottom, y + 1);
                }
            }

            return left == int.MaxValue ? null : new SKRectI(left, top, right, bottom);
        }

        private static int InkedRowsInColumn(SKBitmap bitmap, int x) =>
            Enumerable.Range(0, bitmap.Height).Count(y => Inked(bitmap, x, y));

        private static void AssertBounds(SKRectI? actual, float left, float top, float right, float bottom, float tolerance = 1.5f)
        {
            Assert.True(actual.HasValue, "何も描かれていない");
            var b = actual!.Value;
            Assert.True(Math.Abs(b.Left - left) <= tolerance, $"Left={b.Left} expected={left}");
            Assert.True(Math.Abs(b.Top - top) <= tolerance, $"Top={b.Top} expected={top}");
            Assert.True(Math.Abs(b.Right - right) <= tolerance, $"Right={b.Right} expected={right}");
            Assert.True(Math.Abs(b.Bottom - bottom) <= tolerance, $"Bottom={b.Bottom} expected={bottom}");
        }

        // -- 種類 -------------------------------------------------

        /// <summary>
        /// 線の太さ10・中サイズ(長さ・幅とも30)の矢印を、先端 (100,100)・右向きに描いたときの外接矩形(表形式)。
        /// 三角形系は先端の後ろ側(x=70〜100)、円・ひし形は先端を中心(x=85〜115)に描く。
        /// </summary>
        [Theory]
        //          type                  left  top    right  bottom
        [InlineData(LineEndType.Triangle, 70f, 85f, 100f, 115f)]
        [InlineData(LineEndType.Stealth, 70f, 85f, 100f, 115f)]
        [InlineData(LineEndType.Oval, 85f, 85f, 115f, 115f)]
        [InlineData(LineEndType.Diamond, 85f, 85f, 115f, 115f)]
        public void 塗りつぶす矢印は種類ごとの位置と大きさで先端に描かれる(
            LineEndType type, float left, float top, float right, float bottom)
        {
            using var bitmap = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(type), SKColors.Black, 10f));

            AssertBounds(InkBounds(bitmap), left, top, right, bottom);
        }

        [Fact]
        public void arrowは開いた矢印として線の太さで描かれる()
        {
            using var bitmap = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(LineEndType.Arrow), SKColors.Black, 10f));

            // 2本の線の太さ(10)ぶん外側へ広がり、先端はマイターで尖る。
            var bounds = InkBounds(bitmap);
            Assert.True(bounds.HasValue);
            Assert.InRange(bounds!.Value.Left, 60, 72);
            Assert.InRange(bounds.Value.Right, 100, 115);
            Assert.True(Inked(bitmap, 98, 100), "先端付近が描かれていない");

            // 内側は塗らない(三角形なら塗られる位置)。
            Assert.False(Inked(bitmap, 78, 100));
        }

        [Fact]
        public void triangleは内側を塗りstealthは後ろに切れ込みがありarrowは内側を塗らない()
        {
            using var triangle = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(LineEndType.Triangle), SKColors.Black, 10f));
            using var stealth = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(LineEndType.Stealth), SKColors.Black, 10f));
            using var arrow = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(LineEndType.Arrow), SKColors.Black, 10f));

            // 中心線上の x=74(底辺 70 と stealth の切れ込みの頂点 82 の間)。
            Assert.True(Inked(triangle, 74, 100));
            Assert.False(Inked(stealth, 74, 100));
            Assert.False(Inked(arrow, 78, 100));

            // 先端の近くはどれも描く。
            Assert.True(Inked(triangle, 96, 100));
            Assert.True(Inked(stealth, 96, 100));
        }

        [Fact]
        public void ovalは楕円でありdiamondはひし形である()
        {
            using var oval = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(LineEndType.Oval), SKColors.Black, 10f));
            using var diamond = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(LineEndType.Diamond), SKColors.Black, 10f));

            // (88.5, 92.5) は半径15の円の内側、ひし形(|dx|+|dy|<=15)の外側。
            Assert.True(Inked(oval, 88, 92));
            Assert.False(Inked(diamond, 88, 92));

            // 中心はどちらも描く。
            Assert.True(Inked(oval, 100, 100));
            Assert.True(Inked(diamond, 100, 100));
        }

        [Fact]
        public void 矢印の向きは方向ベクトルに従う()
        {
            using var left = Render(c => LineEndRenderer.Draw(c, Tip, new SKPoint(-3, 0), Medium(LineEndType.Triangle), SKColors.Black, 10f));
            using var down = Render(c => LineEndRenderer.Draw(c, Tip, new SKPoint(0, 2), Medium(LineEndType.Triangle), SKColors.Black, 10f));

            // 方向ベクトルの長さは正規化される。
            AssertBounds(InkBounds(left), 100, 85, 130, 115);
            AssertBounds(InkBounds(down), 85, 70, 115, 100);
        }

        [Fact]
        public void 斜め向きの楕円も方向に合わせて回転する()
        {
            // 幅と長さが違う楕円(幅=小10、長さ=大25)を下向きに描くと、縦長になる。
            using var bitmap = Render(c => LineEndRenderer.Draw(
                c, Tip, new SKPoint(0, 1), new LineEndStyle(LineEndType.Oval, LineEndSize.Small, LineEndSize.Large), SKColors.Black, 5f));

            AssertBounds(InkBounds(bitmap), 95, 87.5f, 105, 112.5f);
        }

        [Fact]
        public void 方向ベクトルの長さが0なら何も描かない()
        {
            using var bitmap = Render(c => LineEndRenderer.Draw(c, Tip, new SKPoint(0, 0), Medium(LineEndType.Triangle), SKColors.Black, 10f));

            Assert.Null(InkBounds(bitmap));
        }

        [Fact]
        public void 指定した色で描く()
        {
            using var bitmap = Render(c => LineEndRenderer.Draw(c, Tip, Right, Medium(LineEndType.Triangle), SKColors.Red, 10f));

            var pixel = bitmap.GetPixel(90, 100);
            Assert.Equal(SKColors.Red, pixel);
        }

        // -- 大きさ -------------------------------------------------

        [Theory]
        [InlineData(LineEndSize.Small, 2f)]
        [InlineData(LineEndSize.Medium, 3f)]
        [InlineData(LineEndSize.Large, 5f)]
        public void 矢印の大きさは線の太さの小2倍中3倍大5倍になる(LineEndSize size, float factor)
        {
            const float lineWidth = 6f;
            using var bitmap = Render(c => LineEndRenderer.Draw(
                c, Tip, Right, new LineEndStyle(LineEndType.Triangle, size, size), SKColors.Black, lineWidth));

            var length = lineWidth * factor;
            AssertBounds(InkBounds(bitmap), 100 - length, 100 - (length / 2), 100, 100 + (length / 2));
        }

        [Fact]
        public void 矢印の幅と長さは独立に決まる()
        {
            // 幅=大(5倍=30)、長さ=小(2倍=12)。
            using var bitmap = Render(c => LineEndRenderer.Draw(
                c, Tip, Right, new LineEndStyle(LineEndType.Triangle, LineEndSize.Large, LineEndSize.Small), SKColors.Black, 6f));

            AssertBounds(InkBounds(bitmap), 88, 85, 100, 115);
        }

        [Theory]
        [InlineData(0.25f)]
        [InlineData(0f)]
        public void 細い線でも倍数を掛ける元の太さは1ptになる(float lineWidth)
        {
            // 大(5倍) × 最小1pt = 5。
            using var bitmap = Render(c =>
            {
                c.Scale(4f, 4f, Tip.X, Tip.Y); // 1pt の大きさを確かめやすいよう拡大する
                LineEndRenderer.Draw(
                    c, Tip, Right, new LineEndStyle(LineEndType.Triangle, LineEndSize.Large, LineEndSize.Large), SKColors.Black, lineWidth);
            });

            AssertBounds(InkBounds(bitmap), 100 - 20, 100 - 10, 100, 100 + 10);
        }

        // -- 線の始点と終点 -------------------------------------------------

        [Fact]
        public void DrawOnPathは始点に外向きのheadEndを描く()
        {
            using var path = new SKPath();
            path.MoveTo(20, 100);
            path.LineTo(180, 100);

            using var bitmap = Render(c => LineEndRenderer.DrawOnPath(c, path, Medium(LineEndType.Triangle), null, SKColors.Black, 10f));

            // 先端 x=20 から線の内側 x=50 まで。先端側が細く、底辺側が太い(外向き)。
            AssertBounds(InkBounds(bitmap), 20, 85, 50, 115);
            Assert.True(InkedRowsInColumn(bitmap, 22) < InkedRowsInColumn(bitmap, 47));
        }

        [Fact]
        public void DrawOnPathは終点に外向きのtailEndを描く()
        {
            using var path = new SKPath();
            path.MoveTo(20, 100);
            path.LineTo(180, 100);

            using var bitmap = Render(c => LineEndRenderer.DrawOnPath(c, path, null, Medium(LineEndType.Triangle), SKColors.Black, 10f));

            AssertBounds(InkBounds(bitmap), 150, 85, 180, 115);
            Assert.True(InkedRowsInColumn(bitmap, 177) < InkedRowsInColumn(bitmap, 152));
        }

        [Fact]
        public void DrawOnPathは折れ線の端の線分の向きに合わせる()
        {
            // 鍵型の接続線(bentConnector)を想定。始点は上向きの線分、終点は右向きの線分。
            using var path = new SKPath();
            path.MoveTo(40, 160);
            path.LineTo(40, 60);
            path.LineTo(160, 60);

            using var bitmap = Render(c => LineEndRenderer.DrawOnPath(
                c, path, Medium(LineEndType.Triangle), Medium(LineEndType.Triangle), SKColors.Black, 10f));

            // 始点(40,160)の矢印は下向き(線の外側)に尖り、y=130〜160。終点(160,60)は右向きで x=130〜160。
            Assert.True(Inked(bitmap, 40, 150));
            Assert.False(Inked(bitmap, 40, 165));
            Assert.True(Inked(bitmap, 150, 60));
            Assert.False(Inked(bitmap, 165, 60));

            var bounds = InkBounds(bitmap)!.Value;
            Assert.InRange(bounds.Bottom, 159, 162);
            Assert.InRange(bounds.Right, 159, 162);
        }

        [Fact]
        public void DrawOnPolylineは始点と終点に外向きの矢印を描く()
        {
            var points = new[] { new SKPoint(20, 100), new SKPoint(100, 100), new SKPoint(100, 180) };

            using var bitmap = Render(c => LineEndRenderer.DrawOnPolyline(
                c, points, Medium(LineEndType.Triangle), Medium(LineEndType.Triangle), SKColors.Black, 10f));

            // 始点は左向き(x=20〜50)、終点は下向き(y=150〜180)。
            Assert.True(InkedRowsInColumn(bitmap, 22) < InkedRowsInColumn(bitmap, 47));
            Assert.True(Inked(bitmap, 100, 175));
            Assert.False(Inked(bitmap, 100, 130));
            var bounds = InkBounds(bitmap)!.Value;
            Assert.InRange(bounds.Left, 18, 22);
            Assert.InRange(bounds.Bottom, 178, 182);
        }

        [Fact]
        public void DrawOnPolylineは端の点と重なる点を飛ばして向きを決める()
        {
            // 回帰テスト(code-reviewer指摘): 以前は先頭2点・末尾2点が重なると向きが0になり矢印が描かれなかった。
            var points = new[]
            {
                new SKPoint(20, 100), new SKPoint(20, 100), new SKPoint(100, 100), new SKPoint(100, 180), new SKPoint(100, 180),
            };

            using var bitmap = Render(c => LineEndRenderer.DrawOnPolyline(
                c, points, Medium(LineEndType.Triangle), Medium(LineEndType.Triangle), SKColors.Black, 10f));

            Assert.True(InkedRowsInColumn(bitmap, 22) < InkedRowsInColumn(bitmap, 47));
            Assert.True(Inked(bitmap, 100, 175));
        }

        [Fact]
        public void 矢印の指定が無ければ何も描かない()
        {
            using var path = new SKPath();
            path.MoveTo(20, 100);
            path.LineTo(180, 100);
            var points = new[] { new SKPoint(20, 100), new SKPoint(180, 100) };

            using var bitmap = Render(c =>
            {
                LineEndRenderer.DrawOnPath(c, path, null, null, SKColors.Black, 10f);
                LineEndRenderer.DrawOnPolyline(c, points, null, null, SKColors.Black, 10f);
            });

            Assert.Null(InkBounds(bitmap));
        }

        [Fact]
        public void 長さ0の線や点が1つしかない折れ線には矢印を描かない()
        {
            using var path = new SKPath();
            path.MoveTo(100, 100);
            path.LineTo(100, 100);

            using var bitmap = Render(c =>
            {
                LineEndRenderer.DrawOnPath(c, path, Medium(LineEndType.Triangle), Medium(LineEndType.Triangle), SKColors.Black, 10f);
                LineEndRenderer.DrawOnPolyline(
                    c, new[] { new SKPoint(100, 100) }, Medium(LineEndType.Triangle), Medium(LineEndType.Triangle), SKColors.Black, 10f);
            });

            Assert.Null(InkBounds(bitmap));
        }
    }
}
