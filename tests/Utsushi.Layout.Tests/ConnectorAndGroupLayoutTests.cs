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
    /// 接続線(<see cref="ConnectorModel"/>/<see cref="GroupChildConnector"/>)とグループ化された
    /// 図形(<see cref="GroupShapeModel"/>/<see cref="GroupChildGroup"/>)のLayoutレイヤーでの
    /// 座標変換の検証(要件10.9, 10.10)。
    /// </summary>
    /// <remarks>
    /// フォントメトリクスは実行環境のフォント構成に依存しないよう
    /// <see cref="ApproximateFontMetricsProvider"/> を使う(<see cref="ReportLayoutEngineTests"/>と同じ方針)。
    /// </remarks>
    public sealed class ConnectorAndGroupLayoutTests
    {
        private readonly ReportLayoutEngine _engine = new(new ApproximateFontMetricsProvider());

        private PagedLayout Compute(SheetModel sheet, ReportDefinition? definition = null) =>
            _engine.Compute(ReportModel.Create(definition ?? Definition(), sheet));

        // -- 要件10.9: 接続線 -------------------------------------------------

        [Fact]
        public void 接続線はアンカーセルの位置とオフセットからページ座標に変換され属性がそのまま渡る()
        {
            const double columnWidthChars = 10.0;
            const double rowHeightPt = 20.0;
            var columnWidthPt = ExcelUnitConverter.ColumnWidthToPoints(columnWidthChars, ReportDefinition.DefaultMaxDigitWidthPx);

            var sheet = UniformSheet(
                rows: 3, columns: 3, columnWidth: columnWidthChars, rowHeightPt: rowHeightPt, pageSetup: NoMarginA4());
            var outline = new ShapeOutline(ArgbColor.Black, 2.5);
            var connector = new ConnectorModel(
                ConnectorPresetType.Bent2Segment,
                45,
                true,
                false,
                outline,
                CellAddress.Parse("B2"),
                new PointPt(2.0, 3.0),
                new FixedAnchorExtent(15.0, 8.0));
            sheet = sheet with { DrawingObjects = new[] { connector } };

            var page = Assert.Single(Compute(sheet).Pages);
            var command = Assert.Single(Connectors(page));

            // ShapeModel/ImageModelと全く同じTryComputeDrawingObjectRectで矩形を求める(design.md参照)。
            Assert.Equal(columnWidthPt + 2.0, command.Rect.Left, 3);
            Assert.Equal(rowHeightPt + 3.0, command.Rect.Top, 3);
            Assert.Equal(15.0, command.Rect.Width, 3);
            Assert.Equal(8.0, command.Rect.Height, 3);

            Assert.Equal(ConnectorPresetType.Bent2Segment, command.Preset);
            Assert.Equal(45, command.RotationDegrees);
            Assert.True(command.FlipHorizontal);
            Assert.False(command.FlipVertical);
            Assert.Equal(outline, command.Outline);
        }

        [Fact]
        public void 改ページをまたぐ接続線はアンカーセルが属するページにのみ配置される()
        {
            var sheet = UniformSheet(
                rows: 4, columns: 2, columnWidth: 10.0, rowHeightPt: 20.0,
                pageSetup: NoMarginA4(rowBreaks: new[] { 3 }));
            var connector = new ConnectorModel(
                ConnectorPresetType.Straight, 0, false, false, null,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(5.0, 5.0));
            sheet = sheet with { DrawingObjects = new[] { connector } };

            var layout = Compute(sheet);
            Assert.Equal(2, layout.PageCount);

            Assert.Single(Connectors(layout.Pages[0]));
            Assert.Empty(Connectors(layout.Pages[1]));
        }

        // -- 要件10.10: グループ -------------------------------------------------

        [Fact]
        public void 単純なグループは子座標空間がグループ自身のサイズと同じ場合子要素が正しい絶対位置になる()
        {
            // ChildOffset=(0,0)、ChildExtent=グループ自身の表示サイズ(=倍率1)のケース。
            const double columnWidthChars = 10.0;
            const double rowHeightPt = 20.0;
            var columnWidthPt = ExcelUnitConverter.ColumnWidthToPoints(columnWidthChars, ReportDefinition.DefaultMaxDigitWidthPx);

            var sheet = UniformSheet(
                rows: 3, columns: 3, columnWidth: columnWidthChars, rowHeightPt: rowHeightPt, pageSetup: NoMarginA4());

            var childShape = new GroupChildShape(
                RectPt.FromBounds(5, 5, 15, 15), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null);
            var group = new GroupShapeModel(
                new PointPt(0, 0),
                new PointPt(40, 40),
                new GroupChildModel[] { childShape },
                0,
                CellAddress.Parse("B2"),
                new PointPt(2.0, 3.0),
                new FixedAnchorExtent(40.0, 40.0));
            sheet = sheet with { DrawingObjects = new[] { group } };

            var page = Assert.Single(Compute(sheet).Pages);
            var groupCommand = Assert.Single(Groups(page));
            var childCommand = Assert.IsType<ShapeCommand>(Assert.Single(groupCommand.Children));

            var groupLeft = columnWidthPt + 2.0;
            var groupTop = rowHeightPt + 3.0;

            Assert.Equal(groupLeft + 5.0, childCommand.Rect.Left, 3);
            Assert.Equal(groupTop + 5.0, childCommand.Rect.Top, 3);
            Assert.Equal(10.0, childCommand.Rect.Width, 3);
            Assert.Equal(10.0, childCommand.Rect.Height, 3);

            // グループ自身の中心・回転もGroupCommandに正しく渡る。
            Assert.Equal(groupLeft + 20.0, groupCommand.Center.X, 3);
            Assert.Equal(groupTop + 20.0, groupCommand.Center.Y, 3);
            Assert.Equal(0, groupCommand.RotationDegrees);
        }

        [Fact]
        public void グループの子座標空間の原点がずれている場合でも平行移動が正しく計算される()
        {
            // ChildOffset != (0,0) の場合、子の位置は「原点からの相対位置」ではなく
            // 「ChildOffsetからの相対位置」でなければならない(平行移動の検証)。
            var sheet = UniformSheet(rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0, pageSetup: NoMarginA4());

            var childShape = new GroupChildShape(
                RectPt.FromBounds(110, 110, 120, 120), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null);
            var group = new GroupShapeModel(
                new PointPt(100, 100),
                new PointPt(40, 40), // グループ自身のサイズと同じ => scale=1
                new GroupChildModel[] { childShape },
                0,
                CellAddress.Parse("A1"),
                new PointPt(0, 0),
                new FixedAnchorExtent(40.0, 40.0));
            sheet = sheet with { DrawingObjects = new[] { group } };

            var page = Assert.Single(Compute(sheet).Pages);
            var groupCommand = Assert.Single(Groups(page));
            var childCommand = Assert.IsType<ShapeCommand>(Assert.Single(groupCommand.Children));

            // グループ自身の矩形左上=(0,0)。子の子座標空間上の位置(110,110)からChildOffset(100,100)を
            // 引いた(10,10)がscale=1のままグループ左上からのオフセットになる(単に群のアンカー位置ぶん
            // ずれるだけの誤りだと(110,110)になってしまうため、平行移動が正しいことの回帰確認になる)。
            Assert.Equal(10.0, childCommand.Rect.Left, 3);
            Assert.Equal(10.0, childCommand.Rect.Top, 3);
            Assert.Equal(10.0, childCommand.Rect.Width, 3);
            Assert.Equal(10.0, childCommand.Rect.Height, 3);
        }

        [Fact]
        public void グループが非一様倍率の場合子要素の幅高さが独立に伸縮する()
        {
            // グループ自身の表示サイズは100x50(幅:高さ=2:1)だが、子座標空間は200x200(正方形)。
            // scaleX = 100/200 = 0.5, scaleY = 50/200 = 0.25 となり、正方形の子要素が
            // 2:1のアスペクト比を持つ矩形に変換されるはず(非一様倍率を許容する設計)。
            var sheet = UniformSheet(rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0, pageSetup: NoMarginA4());

            var childShape = new GroupChildShape(
                RectPt.FromBounds(0, 0, 40, 40), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null);
            var group = new GroupShapeModel(
                new PointPt(0, 0),
                new PointPt(200, 200),
                new GroupChildModel[] { childShape },
                0,
                CellAddress.Parse("A1"),
                new PointPt(0, 0),
                new FixedAnchorExtent(100.0, 50.0));
            sheet = sheet with { DrawingObjects = new[] { group } };

            var page = Assert.Single(Compute(sheet).Pages);
            var groupCommand = Assert.Single(Groups(page));
            var childCommand = Assert.IsType<ShapeCommand>(Assert.Single(groupCommand.Children));

            Assert.Equal(20.0, childCommand.Rect.Width, 3); // 40 * scaleX(0.5)
            Assert.Equal(10.0, childCommand.Rect.Height, 3); // 40 * scaleY(0.25)
            Assert.NotEqual(childCommand.Rect.Width, childCommand.Rect.Height);
        }

        [Fact]
        public void 二段階の入れ子グループは最終的なページ位置が手計算通りになる()
        {
            // 手計算(layout-fidelity-reviewerの例と同じ):
            // 外側グループのページ矩形 = (0,0)-(50,50)。ChildOffset=(0,0), ChildExtent=(100,100)
            //   => scaleX=scaleY = 50/100 = 0.5
            // 入れ子グループの外側子座標空間上の矩形 = (20,20)-(40,40)
            //   => ページ矩形 = (0 + (20-0)*0.5, 0 + (20-0)*0.5)-(0 + (40-0)*0.5, 0 + (40-0)*0.5)
            //               = (10,10)-(20,20) (幅・高さとも10)
            // 入れ子グループ自身: ChildOffset=(0,0), ChildExtent=(200,200)
            //   => scaleX=scaleY = 10/200 = 0.05
            // シェイプの入れ子子座標空間上の矩形 = (50,50)-(150,150)
            //   => ページ矩形 = (10 + 50*0.05, 10 + 50*0.05)-(10 + 150*0.05, 10 + 150*0.05)
            //               = (12.5,12.5)-(17.5,17.5)
            var sheet = UniformSheet(rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0, pageSetup: NoMarginA4());

            var innerShape = new GroupChildShape(
                RectPt.FromBounds(50, 50, 150, 150), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null);
            var nestedGroup = new GroupChildGroup(
                RectPt.FromBounds(20, 20, 40, 40),
                0,
                new PointPt(0, 0),
                new PointPt(200, 200),
                new GroupChildModel[] { innerShape });
            var outerGroup = new GroupShapeModel(
                new PointPt(0, 0),
                new PointPt(100, 100),
                new GroupChildModel[] { nestedGroup },
                0,
                CellAddress.Parse("A1"),
                new PointPt(0, 0),
                new FixedAnchorExtent(50.0, 50.0));
            sheet = sheet with { DrawingObjects = new[] { outerGroup } };

            var page = Assert.Single(Compute(sheet).Pages);
            var outerCommand = Assert.Single(Groups(page));
            var nestedCommand = Assert.IsType<GroupCommand>(Assert.Single(outerCommand.Children));
            var shapeCommand = Assert.IsType<ShapeCommand>(Assert.Single(nestedCommand.Children));

            Assert.Equal(12.5, shapeCommand.Rect.Left, 3);
            Assert.Equal(12.5, shapeCommand.Rect.Top, 3);
            Assert.Equal(17.5, shapeCommand.Rect.Right, 3);
            Assert.Equal(17.5, shapeCommand.Rect.Bottom, 3);
        }

        [Fact]
        public void グループの子座標空間の大きさが0以下の場合は子要素を描画しない()
        {
            // 壊れたジオメトリ(ChildExtent.X<=0)に対する安全弁。例外にならず、子要素が
            // 静かに空になることを確認する(BuildGroupChildren)。
            var sheet = UniformSheet(rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0, pageSetup: NoMarginA4());

            var childShape = new GroupChildShape(
                RectPt.FromBounds(0, 0, 10, 10), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null);
            var group = new GroupShapeModel(
                new PointPt(0, 0),
                new PointPt(0, 100), // X<=0
                new GroupChildModel[] { childShape },
                0,
                CellAddress.Parse("A1"),
                new PointPt(0, 0),
                new FixedAnchorExtent(50.0, 50.0));
            sheet = sheet with { DrawingObjects = new[] { group } };

            var exception = Record.Exception(() => Compute(sheet));
            Assert.Null(exception);

            var page = Assert.Single(Compute(sheet).Pages);
            var groupCommand = Assert.Single(Groups(page));
            Assert.Empty(groupCommand.Children);
        }

        [Fact]
        public void グループの子座標空間の高さが0以下の場合も子要素を描画しない()
        {
            var sheet = UniformSheet(rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0, pageSetup: NoMarginA4());

            var childShape = new GroupChildShape(
                RectPt.FromBounds(0, 0, 10, 10), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null);
            var group = new GroupShapeModel(
                new PointPt(0, 0),
                new PointPt(100, -1), // Y<=0
                new GroupChildModel[] { childShape },
                0,
                CellAddress.Parse("A1"),
                new PointPt(0, 0),
                new FixedAnchorExtent(50.0, 50.0));
            sheet = sheet with { DrawingObjects = new[] { group } };

            var page = Assert.Single(Compute(sheet).Pages);
            var groupCommand = Assert.Single(Groups(page));
            Assert.Empty(groupCommand.Children);
        }

        [Fact]
        public void グループ内の接続線もConnectorCommandになり属性がそのまま渡る()
        {
            var sheet = UniformSheet(rows: 3, columns: 3, columnWidth: 10.0, rowHeightPt: 20.0, pageSetup: NoMarginA4());
            var outline = new ShapeOutline(ArgbColor.Black, 3.0);
            var childConnector = new GroupChildConnector(
                RectPt.FromBounds(0, 0, 40, 40), ConnectorPresetType.Curved3Segment, 10, true, true, outline);
            var group = new GroupShapeModel(
                new PointPt(0, 0),
                new PointPt(40, 40),
                new GroupChildModel[] { childConnector },
                0,
                CellAddress.Parse("A1"),
                new PointPt(0, 0),
                new FixedAnchorExtent(40.0, 40.0));
            sheet = sheet with { DrawingObjects = new[] { group } };

            var page = Assert.Single(Compute(sheet).Pages);
            var groupCommand = Assert.Single(Groups(page));
            var connectorCommand = Assert.IsType<ConnectorCommand>(Assert.Single(groupCommand.Children));

            Assert.Equal(ConnectorPresetType.Curved3Segment, connectorCommand.Preset);
            Assert.Equal(10, connectorCommand.RotationDegrees);
            Assert.True(connectorCommand.FlipHorizontal);
            Assert.True(connectorCommand.FlipVertical);
            Assert.Equal(outline, connectorCommand.Outline);
        }

        [Fact]
        public void 改ページをまたぐグループはアンカーセルが属するページにのみ配置される()
        {
            var sheet = UniformSheet(
                rows: 4, columns: 2, columnWidth: 10.0, rowHeightPt: 20.0,
                pageSetup: NoMarginA4(rowBreaks: new[] { 3 }));
            var childShape = new GroupChildShape(
                RectPt.FromBounds(0, 0, 5, 5), ShapePresetType.Rect, Array.Empty<double>(), 0, null, null, null);
            var group = new GroupShapeModel(
                new PointPt(0, 0),
                new PointPt(5, 5),
                new GroupChildModel[] { childShape },
                0,
                CellAddress.Parse("A1"), default, new FixedAnchorExtent(5.0, 5.0));
            sheet = sheet with { DrawingObjects = new[] { group } };

            var layout = Compute(sheet);
            Assert.Equal(2, layout.PageCount);

            Assert.Single(Groups(layout.Pages[0]));
            Assert.Empty(Groups(layout.Pages[1]));
        }
    }
}
