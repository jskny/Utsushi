using System;
using System.Linq;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using static Utsushi.Layout.Tests.LayoutFixtures;

namespace Utsushi.Layout.Tests
{
    /// <summary>
    /// 図形・グループの反転(要件10.17。design.md「反転」)のLayoutレイヤーでの扱いの検証。
    /// グループの反転は、子の矩形をグループ矩形内で鏡映し、子の反転を XOR で切り替え、片方向だけの反転なら
    /// 子の回転角の符号を反転することで子要素に畳み込む。反転した図形への接続点は矩形の中心で鏡映する。
    /// </summary>
    /// <remarks>
    /// 図形はすべて A1 起点・余白ゼロの用紙に置くため、グループの矩形は (0,0)-(40,40)(子座標空間と同じ大きさ = 倍率1)。
    /// </remarks>
    public sealed class ShapeFlipLayoutTests
    {
        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

        private PageLayout ComputeSinglePage(params DrawingObjectModel[] drawingObjects)
        {
            var sheet = UniformSheet(rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0, pageSetup: NoMarginA4());
            sheet = sheet with { DrawingObjects = drawingObjects };
            return Assert.Single(_engine.Compute(ReportModel.Create(Definition(), sheet)).Pages);
        }

        private static GroupShapeModel Group(bool flipH, bool flipV, params GroupChildModel[] children) =>
            new(
                1u,
                new PointPt(0, 0),
                new PointPt(40, 40),
                children,
                0,
                CellAddress.Parse("A1"),
                new PointPt(0, 0),
                new FixedAnchorExtent(40.0, 40.0),
                flipH,
                flipV);

        private static GroupChildShape ChildShape(
            RectPt localRect, double rotation = 0, bool flipH = false, bool flipV = false, uint id = 7u) =>
            new(id, localRect, ShapePresetType.Rect, Array.Empty<double>(), rotation, null, null, null, flipH, flipV);

        private static void AssertRect(RectPt actual, double left, double top, double right, double bottom)
        {
            Assert.Equal(left, actual.Left, 3);
            Assert.Equal(top, actual.Top, 3);
            Assert.Equal(right, actual.Right, 3);
            Assert.Equal(bottom, actual.Bottom, 3);
        }

        // -- トップレベルの図形 -------------------------------------------------

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void 図形の反転はShapeCommandの反転になり矩形と回転角は変わらない(bool flipH, bool flipV)
        {
            var shape = new ShapeModel(
                2u, ShapePresetType.RightArrow, Array.Empty<double>(), 30, null, null, null,
                CellAddress.Parse("A1"), new PointPt(5, 6), new FixedAnchorExtent(40.0, 20.0), flipH, flipV);

            var command = Assert.Single(Shapes(ComputeSinglePage(shape)));

            Assert.Equal(flipH, command.FlipHorizontal);
            Assert.Equal(flipV, command.FlipVertical);
            Assert.Equal(30, command.RotationDegrees, 6);
            AssertRect(command.Rect, 5, 6, 45, 26);
        }

        // -- グループの反転 -------------------------------------------------

        /// <summary>
        /// 子 (5,5)-(15,20)・回転30度を、グループの反転の組み合わせごとに配置した結果(表形式)。
        /// </summary>
        [Theory]
        //          groupH groupV childH childV  left  top  right bottom rotation flipH  flipV
        [InlineData(false, false, false, false, 5.0, 5.0, 15.0, 20.0, 30.0, false, false)]
        [InlineData(true, false, false, false, 25.0, 5.0, 35.0, 20.0, -30.0, true, false)]
        [InlineData(false, true, false, false, 5.0, 20.0, 15.0, 35.0, -30.0, false, true)]
        [InlineData(true, true, false, false, 25.0, 20.0, 35.0, 35.0, 30.0, true, true)]
        [InlineData(true, false, true, false, 25.0, 5.0, 35.0, 20.0, -30.0, false, false)] // 子自身の反転と打ち消し合う
        [InlineData(true, false, false, true, 25.0, 5.0, 35.0, 20.0, -30.0, true, true)]
        [InlineData(true, true, true, true, 25.0, 20.0, 35.0, 35.0, 30.0, false, false)]
        public void 反転したグループの子は矩形が鏡映され反転が切り替わり片方向の反転では回転角の符号が反転する(
            bool groupH,
            bool groupV,
            bool childH,
            bool childV,
            double left,
            double top,
            double right,
            double bottom,
            double rotation,
            bool expectedH,
            bool expectedV)
        {
            var group = Group(groupH, groupV, ChildShape(RectPt.FromBounds(5, 5, 15, 20), 30, childH, childV));

            var groupCommand = Assert.Single(Groups(ComputeSinglePage(group)));
            var child = Assert.IsType<ShapeCommand>(Assert.Single(groupCommand.Children));

            AssertRect(child.Rect, left, top, right, bottom);
            Assert.Equal(rotation, child.RotationDegrees, 6);
            Assert.Equal(expectedH, child.FlipHorizontal);
            Assert.Equal(expectedV, child.FlipVertical);

            // グループ自身の中心・回転角は反転で変わらない(反転は子要素に畳み込み済み)。
            Assert.Equal(20.0, groupCommand.Center.X, 3);
            Assert.Equal(20.0, groupCommand.Center.Y, 3);
            Assert.Equal(0, groupCommand.RotationDegrees);
        }

        [Fact]
        public void 左右反転したグループ内の画像と接続線も矩形が鏡映され回転角の符号が反転する()
        {
            var image = new GroupChildImage(8u, RectPt.FromBounds(0, 0, 10, 10), Array.Empty<byte>(), "image/png", 20);
            var connector = new GroupChildConnector(
                RectPt.FromBounds(10, 30, 30, 40), ConnectorPresetType.Straight, 45, false, true, null, null, null);
            var group = Group(true, false, image, connector);

            var groupCommand = Assert.Single(Groups(ComputeSinglePage(group)));
            var imageCommand = Assert.IsType<ImageCommand>(groupCommand.Children[0]);
            var connectorCommand = Assert.IsType<ConnectorCommand>(groupCommand.Children[1]);

            AssertRect(imageCommand.Rect, 30, 0, 40, 10);
            Assert.Equal(-20, imageCommand.RotationDegrees, 6);

            AssertRect(connectorCommand.Rect, 10, 30, 30, 40); // 中央に対称な位置なので変わらない
            Assert.Equal(-45, connectorCommand.RotationDegrees, 6);
            Assert.True(connectorCommand.FlipHorizontal);
            Assert.True(connectorCommand.FlipVertical);
        }

        /// <summary>
        /// 外側グループ(左右反転)の中に、(0,0)-(20,20) の入れ子グループ(回転10度)と、その中の子 (0,0)-(5,5)・回転30度。
        /// 入れ子グループも左右反転していれば、子の鏡映は2回で打ち消し合う。
        /// </summary>
        [Theory]
        //          innerH  childLeft childRight childRotation childFlipH
        [InlineData(false, 35.0, 40.0, -30.0, true)]
        [InlineData(true, 20.0, 25.0, 30.0, false)]
        public void 入れ子のグループでは反転が重なる(
            bool innerFlipH, double childLeft, double childRight, double childRotation, bool childFlipH)
        {
            var nested = new GroupChildGroup(
                9u,
                RectPt.FromBounds(0, 0, 20, 20),
                10,
                new PointPt(0, 0),
                new PointPt(20, 20),
                new GroupChildModel[] { ChildShape(RectPt.FromBounds(0, 0, 5, 5), 30) },
                innerFlipH,
                false);
            var group = Group(true, false, nested);

            var outer = Assert.Single(Groups(ComputeSinglePage(group)));
            var nestedCommand = Assert.IsType<GroupCommand>(Assert.Single(outer.Children));
            var child = Assert.IsType<ShapeCommand>(Assert.Single(nestedCommand.Children));

            // 入れ子グループ自身は外側の反転で (20,0)-(40,20) に移り、回転角の符号が反転する。
            Assert.Equal(30.0, nestedCommand.Center.X, 3);
            Assert.Equal(10.0, nestedCommand.Center.Y, 3);
            Assert.Equal(-10, nestedCommand.RotationDegrees, 6);

            AssertRect(child.Rect, childLeft, 0, childRight, 5);
            Assert.Equal(childRotation, child.RotationDegrees, 6);
            Assert.Equal(childFlipH, child.FlipHorizontal);
            Assert.False(child.FlipVertical);
        }

        // -- 反転した図形への接続点 -------------------------------------------------

        /// <summary>矩形 (10,20)-(50,50) の接続点(表形式)。</summary>
        [Theory]
        //          site flipH  flipV  x     y
        [InlineData(0u, false, false, 30.0, 20.0)]
        [InlineData(1u, false, false, 10.0, 35.0)]
        [InlineData(1u, true, false, 50.0, 35.0)] // 左 → 右
        [InlineData(3u, true, false, 10.0, 35.0)] // 右 → 左
        [InlineData(0u, true, false, 30.0, 20.0)] // 上は左右反転で変わらない
        [InlineData(0u, false, true, 30.0, 50.0)] // 上 → 下
        [InlineData(2u, false, true, 30.0, 20.0)] // 下 → 上
        [InlineData(1u, false, true, 10.0, 35.0)] // 左は上下反転で変わらない
        [InlineData(0u, true, true, 30.0, 50.0)]
        [InlineData(1u, true, true, 50.0, 35.0)]
        public void ConnectionSiteResolverは反転した図形の接続点を矩形の中心で鏡映する(
            uint site, bool flipH, bool flipV, double x, double y)
        {
            var point = ConnectionSiteResolver.Resolve(RectPt.FromBounds(10, 20, 50, 50), ShapePresetType.Rect, site, flipH, flipV);

            Assert.Equal(x, point.X, 6);
            Assert.Equal(y, point.Y, 6);
        }

        [Fact]
        public void ConnectionSiteResolverは形状による補正の後で鏡映する()
        {
            // flowChartInputOutput の左の接続点は輪郭に合わせて内側へ補正される(10 + 40*0.2/2 = 14)。
            // 左右反転すると、補正後の点が右側(50 - 4 = 46)へ移る。
            var rect = RectPt.FromBounds(10, 20, 50, 50);

            var unflipped = ConnectionSiteResolver.Resolve(rect, ShapePresetType.FlowChartInputOutput, 1);
            var flipped = ConnectionSiteResolver.Resolve(rect, ShapePresetType.FlowChartInputOutput, 1, flipHorizontal: true);

            Assert.Equal(14.0, unflipped.X, 6);
            Assert.Equal(46.0, flipped.X, 6);
            Assert.Equal(35.0, flipped.Y, 6);
        }

        [Fact]
        public void ConnectionSiteResolverの反転の既定値は反転なし()
        {
            var rect = RectPt.FromBounds(10, 20, 50, 50);

            Assert.Equal(
                ConnectionSiteResolver.Resolve(rect, ShapePresetType.Rect, 1, false, false),
                ConnectionSiteResolver.Resolve(rect, ShapePresetType.Rect, 1));
        }

        [Theory]
        [InlineData(false, false, 0.0, 10.0)]
        [InlineData(true, false, 40.0, 10.0)]
        [InlineData(true, true, 40.0, 10.0)]
        public void 反転した図形を参照する接続線の接続点は鏡映される(bool flipH, bool flipV, double x, double y)
        {
            var target = new ShapeModel(
                5u, ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(40.0, 20.0), flipH, flipV);
            var connector = new ConnectorModel(
                ConnectorPresetType.Straight, 0, false, false, null,
                new ConnectionRef(5u, 1), // 左(idx1)
                null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(5.0, 5.0));

            var command = Assert.Single(Connectors(ComputeSinglePage(target, connector)));

            Assert.NotNull(command.ResolvedStart);
            Assert.Equal(x, command.ResolvedStart!.Value.X, 3);
            Assert.Equal(y, command.ResolvedStart.Value.Y, 3);
        }

        [Theory]
        //          groupH childH  x     y
        [InlineData(false, false, 5.0, 10.0)]  // 子 (5,5)-(15,15) の左
        [InlineData(true, false, 35.0, 10.0)]  // 子は (25,5)-(35,15) に移り、左の接続点は反転で右端へ
        [InlineData(true, true, 25.0, 10.0)]   // 子自身の反転と打ち消し合い、左端のまま
        [InlineData(false, true, 15.0, 10.0)]
        public void 反転したグループ内の図形を参照する接続点は矩形の鏡映と反転の両方を反映する(
            bool groupH, bool childH, double x, double y)
        {
            var group = Group(groupH, false, ChildShape(RectPt.FromBounds(5, 5, 15, 15), flipH: childH, id: 7u));
            var connector = new ConnectorModel(
                ConnectorPresetType.Straight, 0, false, false, null,
                new ConnectionRef(7u, 1), // 左(idx1)
                null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(5.0, 5.0));

            var command = Assert.Single(Connectors(ComputeSinglePage(group, connector)));

            Assert.NotNull(command.ResolvedStart);
            Assert.Equal(x, command.ResolvedStart!.Value.X, 3);
            Assert.Equal(y, command.ResolvedStart.Value.Y, 3);
        }

        [Theory]
        [InlineData(false, 0.0)]  // グループ (0,0)-(40,40) の左の接続点
        [InlineData(true, 40.0)]  // 左右反転したグループの左の接続点は右端へ
        public void 反転したグループ自身を参照する接続点は鏡映される(bool groupH, double x)
        {
            // 回帰テスト: 以前はグループ自身を接続先にした場合に反転を無視していた(test-writer指摘)。
            var group = Group(groupH, false, ChildShape(RectPt.FromBounds(5, 5, 15, 15)));
            var connector = new ConnectorModel(
                ConnectorPresetType.Straight, 0, false, false, null,
                new ConnectionRef(1u, 1), null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(5.0, 5.0));

            var command = Assert.Single(Connectors(ComputeSinglePage(group, connector)));

            Assert.NotNull(command.ResolvedStart);
            Assert.Equal(x, command.ResolvedStart!.Value.X, 3);
            Assert.Equal(20.0, command.ResolvedStart.Value.Y, 3);
        }

        [Theory]
        [InlineData(false, 5.0)]   // 画像 (5,5)-(15,15) の左
        [InlineData(true, 35.0)]   // 画像は (25,5)-(35,15) に移り、左の接続点は鏡映で右端へ
        public void 反転したグループ内の画像を参照する接続点は鏡映される(bool groupH, double x)
        {
            var image = new GroupChildImage(8u, RectPt.FromBounds(5, 5, 15, 15), Array.Empty<byte>(), "image/png", 0);
            var connector = new ConnectorModel(
                ConnectorPresetType.Straight, 0, false, false, null,
                new ConnectionRef(8u, 1), null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(5.0, 5.0));

            var command = Assert.Single(Connectors(ComputeSinglePage(Group(groupH, false, image), connector)));

            Assert.NotNull(command.ResolvedStart);
            Assert.Equal(x, command.ResolvedStart!.Value.X, 3);
            Assert.Equal(10.0, command.ResolvedStart.Value.Y, 3);
        }

        [Fact]
        public void 入れ子のグループ内の図形を参照する接続点も反転の重なりを反映する()
        {
            // 外側(左右反転)→ 入れ子グループ (0,0)-(20,20) は (20,0)-(40,20) へ。
            // 子 (0,0)-(5,5) は入れ子の中でさらに鏡映され (35,0)-(40,5)。左の接続点は反転で右端 (40, 2.5)。
            var nested = new GroupChildGroup(
                9u, RectPt.FromBounds(0, 0, 20, 20), 0, new PointPt(0, 0), new PointPt(20, 20),
                new GroupChildModel[] { ChildShape(RectPt.FromBounds(0, 0, 5, 5), id: 7u) });
            var connector = new ConnectorModel(
                ConnectorPresetType.Straight, 0, false, false, null,
                new ConnectionRef(7u, 1), null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(5.0, 5.0));

            var command = Assert.Single(Connectors(ComputeSinglePage(Group(true, false, nested), connector)));

            Assert.NotNull(command.ResolvedStart);
            Assert.Equal(40.0, command.ResolvedStart!.Value.X, 3);
            Assert.Equal(2.5, command.ResolvedStart.Value.Y, 3);
        }

        [Fact]
        public void 反転したグループ内の図形のテキストは反転しても行の内容と数は変わらない()
        {
            // 文字は反転しない(Rendering で形状にだけ反転をかける)ため、Layout はテキスト行をそのまま作る。
            var font = new FontStyle("Calibri", 10.0, false, false, UnderlineStyle.None, false, ArgbColor.Black);
            var text = new ShapeTextBody(
                new[] { new ShapeTextParagraph(new[] { new ShapeTextRun("AB", font) }, HorizontalAlignment.Left) },
                VerticalAlignment.Top);
            ShapeCommand Build(bool flip)
            {
                var child = new GroupChildShape(
                    7u, RectPt.FromBounds(0, 0, 40, 40), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, text);
                var groupCommand = Assert.Single(Groups(ComputeSinglePage(Group(flip, flip, child))));
                return Assert.IsType<ShapeCommand>(Assert.Single(groupCommand.Children));
            }

            var unflipped = Build(false);
            var flipped = Build(true);

            Assert.True(flipped.FlipHorizontal);
            Assert.True(flipped.FlipVertical);
            Assert.Equal(unflipped.TextLines.Select(l => (l.Origin, l.Text)), flipped.TextLines.Select(l => (l.Origin, l.Text)));
        }
    }
}
