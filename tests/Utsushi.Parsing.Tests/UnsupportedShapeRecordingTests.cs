using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 出力しない図形・接続線・グループを未対応要素として記録すること(excel2pdf-consolidation 要件2.5・2.7(b)(c)(e)、design.md 9)の検証。
    /// </summary>
    public sealed class UnsupportedShapeRecordingTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        [Fact]
        public void 非対応プリセットの図形をプリセット名とアンカーとともに記録する()
        {
            var path = ShapeWorkbookFixtures.CreateWorkbook(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Star10, row: 4, column: 3));
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.UnsupportedShapePreset, element.Kind);
                Assert.Equal(CellAddress.Parse("C4"), element.Anchor);
                Assert.Equal("Star10", element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 非対応プリセットの接続線を記録する()
        {
            var path = ShapeWorkbookFixtures.CreateWorkbook(ConnectorAndGroupWorkbookFixtures.ConnectorAnchor(A.ShapeTypeValues.BentConnector4));
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.UnsupportedShapePreset, element.Kind);
                Assert.Equal("BentConnector4", element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 文字数が上限を超える図形を要素ごとに記録する()
        {
            var longText = new string('あ', 2001);
            var path = ShapeWorkbookFixtures.CreateWorkbook(
                ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, textBody: ShapeWorkbookFixtures.TextBody(longText)),
                ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, row: 10, textBody: ShapeWorkbookFixtures.TextBody(longText), id: 3U));
            try
            {
                var elements = _reader.ReadFile(path).UnsupportedElements;

                Assert.Equal(2, elements.Count);
                Assert.All(elements, e => Assert.Equal(UnsupportedElementKind.LimitExceeded, e.Kind));
                Assert.All(elements, e => Assert.StartsWith("ShapeTextTooLong", e.Detail));
                Assert.All(elements, e => Assert.DoesNotContain("あ", e.Message + e.Detail));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図形の数の上限の超過はシートごとに1件だけ記録する()
        {
            var anchors = Enumerable.Range(0, OpenXmlWorkbookReader.MaxShapesPerSheet + 3)
                .Select(i => (DocumentFormat.OpenXml.OpenXmlElement)ShapeWorkbookFixtures.ShapeAnchor(
                    A.ShapeTypeValues.Rectangle, row: i + 1, id: (uint)(i + 2)))
                .ToArray();
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchors);
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.LimitExceeded, element.Kind);
                Assert.Equal("TooManyShapes: 上限 " + OpenXmlWorkbookReader.MaxShapesPerSheet, element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 非対応の子要素を含むグループはグループのアンカーで1件だけ記録する()
        {
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(new DocumentFormat.OpenXml.OpenXmlElement[]
            {
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Star10, 0, 0, 900000, 900000),
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Star12, 900000, 0, 900000, 900000, id: 11U),
            });
            var path = ShapeWorkbookFixtures.CreateWorkbook(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group, row: 5, column: 4));
            try
            {
                var workbook = _reader.ReadFile(path);

                Assert.Empty(workbook.Sheets[0].DrawingObjects);
                var element = Assert.Single(workbook.UnsupportedElements);
                Assert.Equal(UnsupportedElementKind.UnsupportedGroupChild, element.Kind);
                Assert.Equal(CellAddress.Parse("D5"), element.Anchor);
                Assert.Equal("UnsupportedShapePreset: Star10", element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図表枠を含むグループは要素の名前を原因として記録する()
        {
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(new DocumentFormat.OpenXml.OpenXmlElement[]
            {
                ConnectorAndGroupWorkbookFixtures.GraphicFrameElement(),
            });
            var path = ShapeWorkbookFixtures.CreateWorkbook(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group));
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.UnsupportedGroupChild, element.Kind);
                Assert.Equal("graphicFrame", element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 入れ子の段数が上限を超えるグループは上限の種類を原因として記録する()
        {
            var leaf = ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Rectangle, 0, 0, 900000, 900000);
            var nested = ConnectorAndGroupWorkbookFixtures.WrapNestedGroups(OpenXmlWorkbookReader.MaxShapeNestingDepth, leaf);
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(new[] { nested });
            var path = ShapeWorkbookFixtures.CreateWorkbook(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group));
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.UnsupportedGroupChild, element.Kind);
                Assert.Equal("GroupNestingTooDeep", element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 対応済みの要素だけのグループは記録しない()
        {
            var group = ConnectorAndGroupWorkbookFixtures.GroupShapeElement(new DocumentFormat.OpenXml.OpenXmlElement[]
            {
                ConnectorAndGroupWorkbookFixtures.GroupChildShapeElement(A.ShapeTypeValues.Rectangle, 0, 0, 900000, 900000),
            });
            var path = ShapeWorkbookFixtures.CreateWorkbook(ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(group));
            try
            {
                var workbook = _reader.ReadFile(path);

                Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<GroupShapeModel>());
                Assert.Empty(workbook.UnsupportedElements);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
