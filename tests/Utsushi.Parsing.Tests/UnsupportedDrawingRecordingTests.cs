using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Parsing.OpenXml;
using Utsushi.TestSupport;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 読み込みの対象でない描画オブジェクト・レガシー図形・ヘッダー/フッターの画像・結合セル数の上限の超過を、
    /// 未対応要素として記録すること(excel2pdf-consolidation 要件2.7(a)(e)(f)(h)、design.md 2.2)の検証。
    /// </summary>
    public sealed class UnsupportedDrawingRecordingTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        [Fact]
        public void グラフを参照する図表枠をグラフとしてアンカーとともに記録する()
        {
            var path = ShapeWorkbookFixtures.CreateWorkbook(
                ConnectorAndGroupWorkbookFixtures.WrapInTwoCellAnchor(ChartFrame(), fromRow: 10, fromColumn: 4));
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.Chart, element.Kind);
                Assert.Equal("テストシート", element.SheetName);
                Assert.Equal(CellAddress.Parse("D10"), element.Anchor);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グラフ以外の図表枠を対応していない描画オブジェクトとして記録する()
        {
            var path = ShapeWorkbookFixtures.CreateWorkbook(
                ConnectorAndGroupWorkbookFixtures.WrapInOneCellAnchor(ConnectorAndGroupWorkbookFixtures.GraphicFrameElement()));
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.UnsupportedDrawingObject, element.Kind);
                Assert.Equal(CellAddress.Parse("B3"), element.Anchor);
                Assert.Equal("graphicFrame", element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 絶対位置で配置された描画オブジェクトをアンカーのセル無しで記録する()
        {
            var shape = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle).GetFirstChild<Xdr.Shape>()!.CloneNode(true);
            var absolute = new Xdr.AbsoluteAnchor(
                new Xdr.Position { X = 0L, Y = 0L },
                new Xdr.Extent { Cx = 900000L, Cy = 900000L },
                shape,
                new Xdr.ClientData());
            var path = ShapeWorkbookFixtures.CreateWorkbook(absolute);
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.UnsupportedAnchor, element.Kind);
                Assert.Null(element.Anchor);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 対応済みの描画オブジェクトだけなら記録しない()
        {
            var path = ShapeWorkbookFixtures.CreateWorkbook(ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle));
            try
            {
                Assert.Empty(_reader.ReadFile(path).UnsupportedElements);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void レガシー図形とヘッダーフッターの画像をシートごとに記録する()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook((_, worksheet) =>
            {
                worksheet.Append(new X.LegacyDrawing { Id = "rIdVml" });
                worksheet.Append(new X.LegacyDrawingHeaderFooter { Id = "rIdVmlHf" });
            });
            try
            {
                var elements = _reader.ReadFile(path).UnsupportedElements;

                Assert.Equal(
                    new[] { UnsupportedElementKind.LegacyDrawing, UnsupportedElementKind.HeaderFooterImage },
                    elements.Select(e => e.Kind).ToArray());
                Assert.All(elements, e => Assert.Null(e.Anchor));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 結合セル範囲の数の上限の超過を1件だけ記録する()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook((_, worksheet) =>
            {
                var mergeCells = new X.MergeCells();
                for (var i = 1; i <= OpenXmlWorkbookReader.MaxMergedRangesPerSheet + 2; i++)
                {
                    mergeCells.Append(new X.MergeCell { Reference = $"A{i * 2}:B{i * 2}" });
                }

                worksheet.InsertAfter(mergeCells, worksheet.GetFirstChild<X.SheetData>());
            });
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.LimitExceeded, element.Kind);
                Assert.Equal("TooManyMergedRanges: 上限 " + OpenXmlWorkbookReader.MaxMergedRangesPerSheet, element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void エッジケースの注文書のグラフとコメントを記録する()
        {
            var path = Path.Combine(TestPaths.EdgeCaseSamplesRoot, "edge-order-form", "template.xlsx");
            using var stream = File.OpenRead(path);

            var workbook = _reader.Read(stream, new WorkbookReadOptions(ActiveSheetOnly: true));

            Assert.Equal("注文書", Assert.Single(workbook.Sheets).Name);
            var chart = Assert.Single(workbook.UnsupportedElements, e => e.Kind == UnsupportedElementKind.Chart);
            Assert.Equal("注文書", chart.SheetName);
            Assert.NotNull(chart.Anchor);
            Assert.Contains(workbook.UnsupportedElements, e => e.Kind == UnsupportedElementKind.LegacyDrawing);
            Assert.DoesNotContain(workbook.UnsupportedElements, e => e.Kind == UnsupportedElementKind.UnsupportedDrawingObject);
        }

        private static Xdr.GraphicFrame ChartFrame() =>
            new(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 5U, Name = "Chart 1" },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L }),
                new A.Graphic(
                    new A.GraphicData(new C.ChartReference { Id = "rIdChart" })
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart",
                    }));
    }
}
