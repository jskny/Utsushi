using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// シート内接続線(<c>xdr:cxnSp</c>。要件10.9)とグループ化された図形
    /// (<c>xdr:grpSp</c>。要件10.10)の読み取りの検証。
    /// </summary>
    public sealed class ConnectorAndGroupReadingTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        public static IEnumerable<object[]> SupportedConnectorPresetCases()
        {
            yield return new object[] { A.ShapeTypeValues.StraightConnector1, ConnectorPresetType.Straight };
            yield return new object[] { A.ShapeTypeValues.Line, ConnectorPresetType.Straight }; // Excelの「直線」
            yield return new object[] { A.ShapeTypeValues.BentConnector2, ConnectorPresetType.Bent2Segment };
            yield return new object[] { A.ShapeTypeValues.BentConnector3, ConnectorPresetType.Bent3Segment };
            yield return new object[] { A.ShapeTypeValues.CurvedConnector2, ConnectorPresetType.Curved2Segment };
            yield return new object[] { A.ShapeTypeValues.CurvedConnector3, ConnectorPresetType.Curved3Segment };
        }

        // --- 接続線(要件10.9) ---

        [Theory]
        [MemberData(nameof(SupportedConnectorPresetCases))]
        public void 対応済み接続線プリセットを回転_反転_枠線とともに読み取れる(
            A.ShapeTypeValues preset, ConnectorPresetType expected)
        {
            // 60,000分の1度単位。45度 = 2,700,000。
            var anchor = ConnectorAndGroupWorkbookFixtures.ConnectorAnchor(
                preset,
                rotationEmu: 2_700_000,
                flipH: true,
                flipV: false,
                outline: ShapeWorkbookFixtures.Outline("FF0000", 25400));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var connector = Assert.Single(sheet.DrawingObjects.OfType<ConnectorModel>());

                Assert.Equal(expected, connector.Preset);
                Assert.Equal(45.0, connector.RotationDegrees, 3);
                Assert.True(connector.FlipHorizontal);
                Assert.False(connector.FlipVertical);
                Assert.Equal(CellAddress.Parse("B3"), connector.AnchorCell);
                Assert.IsType<FixedAnchorExtent>(connector.Extent);
                Assert.NotNull(connector.Outline);
                Assert.Equal(new ArgbColor(0xFF, 0xFF, 0x00, 0x00), connector.Outline!.Color);
                Assert.Equal(2.0, connector.Outline.WidthPt, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 非対応の接続線プリセットはunsupportedElementsがignoreなら無視される()
        {
            var anchor = ConnectorAndGroupWorkbookFixtures.ConnectorAnchor(A.ShapeTypeValues.BentConnector4);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects.OfType<ConnectorModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 非対応の接続線プリセットはunsupportedElementsがerrorなら例外になる()
        {
            var anchor = ConnectorAndGroupWorkbookFixtures.ConnectorAnchor(A.ShapeTypeValues.BentConnector4);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("UnsupportedShapePreset", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 接続線のみのシートはunsupportedElementsがerrorでも例外にならない()
        {
            var anchor = ConnectorAndGroupWorkbookFixtures.ConnectorAnchor(A.ShapeTypeValues.StraightConnector1);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<ConnectorModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 接続線のstCxnとendCxnを接続先として読み取れる()
        {
            // 要件10.11: a:stCxn/a:endCxnをConnectionRef(ShapeId, SiteIndex)として読み取る。
            var anchor = ConnectorAndGroupWorkbookFixtures.ConnectorAnchor(
                A.ShapeTypeValues.StraightConnector1,
                startConnection: (5U, 0U),
                endConnection: (7U, 3U));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var connector = Assert.Single(sheet.DrawingObjects.OfType<ConnectorModel>());

                Assert.Equal(new ConnectionRef(5U, 0U), connector.StartConnection);
                Assert.Equal(new ConnectionRef(7U, 3U), connector.EndConnection);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 接続線にstCxnとendCxnが無い場合はnullになる()
        {
            var anchor = ConnectorAndGroupWorkbookFixtures.ConnectorAnchor(A.ShapeTypeValues.StraightConnector1);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var connector = Assert.Single(sheet.DrawingObjects.OfType<ConnectorModel>());

                Assert.Null(connector.StartConnection);
                Assert.Null(connector.EndConnection);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内接続線のstCxnとendCxnも接続先として読み取れる()
        {
            var connector = ConnectorAndGroupWorkbookFixtures.GroupChildConnectorElement(
                A.ShapeTypeValues.StraightConnector1, 0L, 0L, 900000L, 900000L,
                id: 12U,
                startConnection: (10U, 1U),
                endConnection: (11U, 2U));
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { connector }, 0L, 0L, 900000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var readGroup = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
                var readConnector = Assert.IsType<GroupChildConnector>(Assert.Single(readGroup.Children));

                Assert.Equal(new ConnectionRef(10U, 1U), readConnector.StartConnection);
                Assert.Equal(new ConnectionRef(11U, 2U), readConnector.EndConnection);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void twoCellAnchorの接続線のみのシートはunsupportedElementsがerrorでも例外にならない()
        {
            var connector = ConnectorAndGroupWorkbookFixtures.GroupChildConnectorElement(
                A.ShapeTypeValues.StraightConnector1, 0L, 0L, 900000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInTwoCellAnchor(connector);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                var readConnector = Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<ConnectorModel>());
                Assert.IsType<CellSpanAnchorExtent>(readConnector.Extent);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- グループ(要件10.10) ---

        [Fact]
        public void 図形_画像_接続線を含むグループを子座標空間の矩形とともに読み取る()
        {
            const long chOffX = 100000L;
            const long chOffY = 200000L;
            const long chExtCx = 3600000L;
            const long chExtCy = 1800000L;

            const long shapeOffX = 0L;
            const long shapeOffY = 0L;
            const long shapeExtCx = 900000L;
            const long shapeExtCy = 900000L;

            const long imageOffX = 1000000L;
            const long imageOffY = 0L;
            const long imageExtCx = 800000L;
            const long imageExtCy = 800000L;

            const long connOffX = 2000000L;
            const long connOffY = 100000L;
            const long connExtCx = 700000L;
            const long connExtCy = 700000L;

            var path = ConnectorAndGroupWorkbookFixtures.CreateWorkbookWithImage((drawingsPart, imagePart) =>
            {
                var shape = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                    A.ShapeTypeValues.Rectangle, shapeOffX, shapeOffY, shapeExtCx, shapeExtCy, id: 10U);
                var image = ConnectorAndGroupWorkbookFixtures.GroupChildImageElement(
                    drawingsPart, imagePart, imageOffX, imageOffY, imageExtCx, imageExtCy, id: 11U);
                var connector = ConnectorAndGroupWorkbookFixtures.GroupChildConnectorElement(
                    A.ShapeTypeValues.StraightConnector1, connOffX, connOffY, connExtCx, connExtCy, id: 12U);

                var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                    new OpenXmlElement[] { shape, image, connector }, chOffX, chOffY, chExtCx, chExtCy, id: 2U);

                return ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group, widthEmu: chExtCx, heightEmu: chExtCy);
            });

            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());

                Assert.Equal(CellAddress.Parse("B3"), group.AnchorCell);
                // 要件10.11: グループ自身・グループ内の図形/画像のIdもNonVisualDrawingProperties/@idを反映する
                // (接続線の接続先解決のキーとして使うため。GroupChildConnectorは接続先として参照される
                // 対象ではないためIdを持たない)。
                Assert.Equal(2U, group.Id);
                Assert.Equal(Units.EmusToPoints(chOffX), group.ChildOffset.X, 3);
                Assert.Equal(Units.EmusToPoints(chOffY), group.ChildOffset.Y, 3);
                Assert.Equal(Units.EmusToPoints(chExtCx), group.ChildExtent.X, 3);
                Assert.Equal(Units.EmusToPoints(chExtCy), group.ChildExtent.Y, 3);
                Assert.Equal(3, group.Children.Count);

                var shapeChild = Assert.IsType<GroupChildShape>(group.Children[0]);
                Assert.Equal(10U, shapeChild.Id);
                Assert.Equal(ShapePresetType.Rect, shapeChild.Preset);
                AssertRect(shapeOffX, shapeOffY, shapeExtCx, shapeExtCy, shapeChild.LocalRect);

                var imageChild = Assert.IsType<GroupChildImage>(group.Children[1]);
                Assert.Equal(11U, imageChild.Id);
                AssertRect(imageOffX, imageOffY, imageExtCx, imageExtCy, imageChild.LocalRect);
                Assert.Equal("image/png", imageChild.ContentType);

                var connectorChild = Assert.IsType<GroupChildConnector>(group.Children[2]);
                Assert.Equal(ConnectorPresetType.Straight, connectorChild.Preset);
                AssertRect(connOffX, connOffY, connExtCx, connExtCy, connectorChild.LocalRect);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内画像の回転角を度に変換して読み取る()
        {
            // 60,000分の1度単位。90度 = 5,400,000(要件9.7。グループ内画像でも図形と同じ変換)。
            var path = ConnectorAndGroupWorkbookFixtures.CreateWorkbookWithImage((drawingsPart, imagePart) =>
            {
                var image = ConnectorAndGroupWorkbookFixtures.GroupChildImageElement(
                    drawingsPart, imagePart, 0L, 0L, 900000L, 900000L, id: 11U, rotationEmu: 5_400_000);
                var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                    new OpenXmlElement[] { image }, 0L, 0L, 900000L, 900000L, id: 2U);

                return ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group, widthEmu: 900000L, heightEmu: 900000L);
            });

            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
                var imageChild = Assert.IsType<GroupChildImage>(Assert.Single(group.Children));

                Assert.Equal(90.0, imageChild.RotationDegrees, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内に非対応プリセットの図形が1つでもあればunsupportedElementsがignoreならグループ全体が破棄される()
        {
            var goodShape1 = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 900000L, 900000L, id: 10U);
            var badShape = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.FlowChartPreparation, 900000L, 0L, 900000L, 900000L, id: 11U);
            var goodShape2 = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Ellipse, 1800000L, 0L, 900000L, 900000L, id: 12U);

            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { goodShape1, badShape, goodShape2 }, 0L, 0L, 2700000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);

                // グループの一部(良品の2つの子)だけが読み取られるのではなく、グループ全体が
                // 描画オブジェクトから消える(要件10.7補足)。
                Assert.Empty(sheet.DrawingObjects);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内に非対応プリセットの図形が1つでもあればunsupportedElementsがerrorなら例外になる()
        {
            var goodShape = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 900000L, 900000L, id: 10U);
            var badShape = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.FlowChartPreparation, 900000L, 0L, 900000L, 900000L, id: 11U);

            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { goodShape, badShape }, 0L, 0L, 1800000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("UnsupportedShapePreset", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内に非対応の描画オブジェクト種別が含まれる場合はignoreならグループ全体が破棄される()
        {
            var goodShape = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 900000L, 900000L, id: 10U);
            var graphicFrame = ConnectorAndGroupWorkbookFixtures.GraphicFrameElement();

            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { goodShape, graphicFrame }, 0L, 0L, 1800000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内に非対応の描画オブジェクト種別が含まれる場合はerrorなら例外になる()
        {
            var goodShape = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 900000L, 900000L, id: 10U);
            var graphicFrame = ConnectorAndGroupWorkbookFixtures.GraphicFrameElement();

            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { goodShape, graphicFrame }, 0L, 0L, 1800000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("UnsupportedShapePreset", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 入れ子グループを子座標空間の矩形とともに読み取る()
        {
            const long innerOffX = 50000L;
            const long innerOffY = 60000L;
            const long innerExtCx = 1000000L;
            const long innerExtCy = 1000000L;
            const long innerChOffX = 0L;
            const long innerChOffY = 0L;
            const long innerChExtCx = 500000L;
            const long innerChExtCy = 500000L;
            const long leafOffX = 0L;
            const long leafOffY = 0L;
            const long leafExtCx = 300000L;
            const long leafExtCy = 300000L;

            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, leafOffX, leafOffY, leafExtCx, leafExtCy, id: 20U);
            var nested = ConnectorAndGroupWorkbookFixtures.NestedGroupElement(
                new OpenXmlElement[] { leaf },
                innerOffX, innerOffY, innerExtCx, innerExtCy,
                innerChOffX, innerChOffY, innerChExtCx, innerChExtCy,
                id: 21U);

            var outerGroup = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { nested }, 0L, 0L, 2000000L, 2000000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(outerGroup);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
                var nestedChild = Assert.IsType<GroupChildGroup>(Assert.Single(group.Children));

                // 要件10.11: 入れ子グループのIdも接続先解決のキーとして保持する。
                Assert.Equal(21U, nestedChild.Id);
                AssertRect(innerOffX, innerOffY, innerExtCx, innerExtCy, nestedChild.LocalRect);
                Assert.Equal(Units.EmusToPoints(innerChOffX), nestedChild.ChildOffset.X, 3);
                Assert.Equal(Units.EmusToPoints(innerChOffY), nestedChild.ChildOffset.Y, 3);
                Assert.Equal(Units.EmusToPoints(innerChExtCx), nestedChild.ChildExtent.X, 3);
                Assert.Equal(Units.EmusToPoints(innerChExtCy), nestedChild.ChildExtent.Y, 3);

                var leafChild = Assert.IsType<GroupChildShape>(Assert.Single(nestedChild.Children));
                Assert.Equal(ShapePresetType.Rect, leafChild.Preset);
                AssertRect(leafOffX, leafOffY, leafExtCx, leafExtCy, leafChild.LocalRect);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループのネストは上限段数のちょうど境界まで読み取れる()
        {
            // トップレベルのグループ自身が1段目のため、上限(5段)まで許容されるネストは
            // トップレベルの直接の子から数えて MaxShapeNestingDepth-1 段(=4段)分。
            var nestedLevels = OpenXmlWorkbookReader.MaxShapeNestingDepth - 1;
            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U);
            var nestedChain = ConnectorAndGroupWorkbookFixtures.WrapNestedGroups(nestedLevels, leaf);
            var topGroup = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new[] { nestedChain }, 0L, 0L, 900000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(topGroup);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                var group = Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<GroupShapeModel>());

                IReadOnlyList<GroupChildModel> current = group.Children;
                for (var i = 0; i < nestedLevels; i++)
                {
                    var nested = Assert.IsType<GroupChildGroup>(Assert.Single(current));
                    current = nested.Children;
                }

                var leafChild = Assert.IsType<GroupChildShape>(Assert.Single(current));
                Assert.Equal(ShapePresetType.Rect, leafChild.Preset);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループのネストが上限を超える場合はunsupportedElementsがignoreならグループ全体が破棄される()
        {
            // トップ(1段目)+5段のネストで6段目に到達し、上限(5段)を超える。
            var nestedLevels = OpenXmlWorkbookReader.MaxShapeNestingDepth;
            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U);
            var nestedChain = ConnectorAndGroupWorkbookFixtures.WrapNestedGroups(nestedLevels, leaf);
            var topGroup = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new[] { nestedChain }, 0L, 0L, 900000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(topGroup);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループのネストが上限を超える場合はunsupportedElementsがerrorなら例外になる()
        {
            var nestedLevels = OpenXmlWorkbookReader.MaxShapeNestingDepth;
            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U);
            var nestedChain = ConnectorAndGroupWorkbookFixtures.WrapNestedGroups(nestedLevels, leaf);
            var topGroup = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new[] { nestedChain }, 0L, 0L, 900000L, 900000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(topGroup);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("GroupNestingTooDeep", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内の図形はトップレベルの図形と上限個数を共有し超過分はignoreならグループ全体が破棄される()
        {
            var topLevelCount = OpenXmlWorkbookReader.MaxShapesPerSheet - 2; // 48
            var anchors = new List<OpenXmlElement>();
            for (var i = 0; i < topLevelCount; i++)
            {
                anchors.Add(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, row: i + 1, id: (uint)(i + 2)));
            }

            var groupChildren = new OpenXmlElement[]
            {
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Ellipse, 300000L, 0L, 300000L, 300000L, id: 901U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Triangle, 600000L, 0L, 300000L, 300000L, id: 902U),
            };
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                groupChildren, 0L, 0L, 900000L, 300000L, id: (uint)(topLevelCount + 10));
            anchors.Add(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group, row: topLevelCount + 1));

            var path = ShapeWorkbookFixtures.CreateWorkbook(anchors.ToArray());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(topLevelCount, sheet.DrawingObjects.OfType<ShapeModel>().Count());
                Assert.Empty(sheet.DrawingObjects.OfType<GroupShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内の図形はトップレベルの図形と上限個数を共有し超過分はerrorなら例外になる()
        {
            var topLevelCount = OpenXmlWorkbookReader.MaxShapesPerSheet - 2; // 48
            var anchors = new List<OpenXmlElement>();
            for (var i = 0; i < topLevelCount; i++)
            {
                anchors.Add(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, row: i + 1, id: (uint)(i + 2)));
            }

            var groupChildren = new OpenXmlElement[]
            {
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Ellipse, 300000L, 0L, 300000L, 300000L, id: 901U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Triangle, 600000L, 0L, 300000L, 300000L, id: 902U),
            };
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                groupChildren, 0L, 0L, 900000L, 300000L, id: (uint)(topLevelCount + 10));
            anchors.Add(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group, row: topLevelCount + 1));

            var path = ShapeWorkbookFixtures.CreateWorkbook(anchors.ToArray());
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("TooManyShapes", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ自身を1件として数える時点で上限に達する場合はignoreならグループ全体が破棄される()
        {
            // 子要素の処理自体はどれも上限に抵触せず成功するが(47→48→49→50)、
            // グループ自身を1件として数える直前の再チェックで上限(50)に達しているため、
            // グループ全体を破棄する。子要素の走査中に失敗する共有カウント上限のテスト
            // (直前のテストケース)とは異なるコードパス(グループ自身の計上時の再チェック)を
            // 検証する。
            var topLevelCount = OpenXmlWorkbookReader.MaxShapesPerSheet - 3; // 47
            var anchors = new List<OpenXmlElement>();
            for (var i = 0; i < topLevelCount; i++)
            {
                anchors.Add(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, row: i + 1, id: (uint)(i + 2)));
            }

            var groupChildren = new OpenXmlElement[]
            {
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Ellipse, 300000L, 0L, 300000L, 300000L, id: 901U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Triangle, 600000L, 0L, 300000L, 300000L, id: 902U),
            };
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                groupChildren, 0L, 0L, 900000L, 300000L, id: (uint)(topLevelCount + 10));
            anchors.Add(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group, row: topLevelCount + 1));

            var path = ShapeWorkbookFixtures.CreateWorkbook(anchors.ToArray());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(topLevelCount, sheet.DrawingObjects.OfType<ShapeModel>().Count());
                Assert.Empty(sheet.DrawingObjects.OfType<GroupShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ自身を1件として数える時点で上限に達する場合はerrorなら例外になる()
        {
            var topLevelCount = OpenXmlWorkbookReader.MaxShapesPerSheet - 3; // 47
            var anchors = new List<OpenXmlElement>();
            for (var i = 0; i < topLevelCount; i++)
            {
                anchors.Add(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, row: i + 1, id: (uint)(i + 2)));
            }

            var groupChildren = new OpenXmlElement[]
            {
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 900U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Ellipse, 300000L, 0L, 300000L, 300000L, id: 901U),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Triangle, 600000L, 0L, 300000L, 300000L, id: 902U),
            };
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                groupChildren, 0L, 0L, 900000L, 300000L, id: (uint)(topLevelCount + 10));
            anchors.Add(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group, row: topLevelCount + 1));

            var path = ShapeWorkbookFixtures.CreateWorkbook(anchors.ToArray());
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("TooManyShapes", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループのみのシートはunsupportedElementsがerrorでも例外にならない()
        {
            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 10U);
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { leaf }, 0L, 0L, 300000L, 300000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                var group2 = Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<GroupShapeModel>());
                Assert.Single(group2.Children);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void twoCellAnchorのグループのみのシートはunsupportedElementsがerrorでも例外にならない()
        {
            // HasUnsupportedDrawingObjectはoneCellAnchor/twoCellAnchor双方の分岐でgrpSpを
            // 構造的にサポート済みとして扱う必要がある(要件10.7)。
            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(
                A.ShapeTypeValues.Rectangle, 0L, 0L, 300000L, 300000L, id: 10U);
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(
                new OpenXmlElement[] { leaf }, 0L, 0L, 300000L, 300000L, id: 2U);
            var anchor = ConnectorAndGroupWorkbookFixtures.WrapInTwoCellAnchor(group);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                var group2 = Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<GroupShapeModel>());
                Assert.IsType<CellSpanAnchorExtent>(group2.Extent);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static void AssertRect(long xEmu, long yEmu, long cxEmu, long cyEmu, RectPt actual)
        {
            Assert.Equal(Units.EmusToPoints(xEmu), actual.Left, 3);
            Assert.Equal(Units.EmusToPoints(yEmu), actual.Top, 3);
            Assert.Equal(Units.EmusToPoints(cxEmu), actual.Width, 3);
            Assert.Equal(Units.EmusToPoints(cyEmu), actual.Height, 3);
        }
    }
}
