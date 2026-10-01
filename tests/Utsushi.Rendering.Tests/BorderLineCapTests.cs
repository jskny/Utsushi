using SkiaSharp;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>罫線の端の形(太い罫線の外側の角が欠けないこと)の検証。</summary>
    public sealed class BorderLineCapTests
    {
        private static LineCommand Line(PointPt from, PointPt to, LineDashStyle dash = LineDashStyle.Solid) =>
            new(from, to, ArgbColor.Black, 2.25, dash);

        [Fact]
        public void 実線の水平垂直の罫線は端を線幅の半分延ばす()
        {
            Assert.Equal(SKStrokeCap.Square, SkiaPdfRenderer.ResolveLineCap(Line(new PointPt(0, 10), new PointPt(50, 10))));
            Assert.Equal(SKStrokeCap.Square, SkiaPdfRenderer.ResolveLineCap(Line(new PointPt(10, 0), new PointPt(10, 50))));
        }

        [Fact]
        public void 破線と斜めの罫線は端を延ばさない()
        {
            Assert.Equal(
                SKStrokeCap.Butt,
                SkiaPdfRenderer.ResolveLineCap(Line(new PointPt(0, 10), new PointPt(50, 10), LineDashStyle.Dash)));
            Assert.Equal(SKStrokeCap.Butt, SkiaPdfRenderer.ResolveLineCap(Line(new PointPt(0, 0), new PointPt(50, 20))));
        }
    }
}
