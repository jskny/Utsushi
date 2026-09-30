using System;
using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// シート内図形(要件10)の読み取りの検証。
    /// </summary>
    public sealed class ShapeReadingTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        [Fact]
        public void 対応済みプリセットの図形を読み取れる()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                A.ShapeTypeValues.Rectangle,
                fill: ShapeWorkbookFixtures.SolidFill("1F4E8C"),
                outline: ShapeWorkbookFixtures.Outline("000000", 12700));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Equal(ShapePresetType.Rect, shape.Preset);
                Assert.Equal(CellAddress.Parse("B3"), shape.AnchorCell);
                Assert.IsType<FixedAnchorExtent>(shape.Extent);
                Assert.IsType<SolidShapeFill>(shape.Fill);
                Assert.Equal(new ArgbColor(0xFF, 0x1F, 0x4E, 0x8C), ((SolidShapeFill)shape.Fill!).Color);
                Assert.NotNull(shape.Outline);
                Assert.Equal(new ArgbColor(0xFF, 0x00, 0x00, 0x00), shape.Outline!.Color);
                Assert.Null(shape.Text);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図形のIdはNonVisualDrawingProperties_idを反映する()
        {
            // 要件10.11: 接続線の接続先解決のキーとなるIdを図形側でも保持する。
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, id: 42U);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Equal(42u, shape.Id);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 雲形吹き出しのadj1とadj2を引き出し位置として読み取る()
        {
            // 要件10.13: cloudCalloutのadj1(X方向)/adj2(Y方向)は引き出し三角形の位置を表す
            // 調整ガイドとして読み取る(輪郭自体は固定形状のまま調整ガイドに対応しない)。
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                A.ShapeTypeValues.CloudCallout,
                adjustValueList: ShapeWorkbookFixtures.AdjustValues(("adj1", 15000), ("adj2", -20000)));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Equal(ShapePresetType.CloudCallout, shape.Preset);
                Assert.Equal(2, shape.AdjustmentValues.Count);
                Assert.Equal(0.15, shape.AdjustmentValues[0], 3);
                Assert.Equal(-0.2, shape.AdjustmentValues[1], 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("Callout1", ShapePresetType.Callout1, 4)]
        [InlineData("Callout3", ShapePresetType.Callout3, 8)]
        [InlineData("BorderCallout1", ShapePresetType.BorderCallout1, 4)]
        [InlineData("BorderCallout2", ShapePresetType.BorderCallout2, 6)]
        [InlineData("BorderCallout3", ShapePresetType.BorderCallout3, 8)]
        [InlineData("AccentCallout2", ShapePresetType.AccentCallout2, 6)]
        [InlineData("AccentBorderCallout2", ShapePresetType.AccentBorderCallout2, 6)]
        public void 線吹き出しを読み取り引き出し線の調整ガイドを折れ数に応じた個数で返す(
            string presetName, ShapePresetType expected, int expectedGuideCount)
        {
            // 要件10.14: 線吹き出しは (adj1=y1, adj2=x1), (adj3=y2, adj4=x2), ... の組で引き出し線の
            // 頂点を持つ。ファイルに無いガイドはNaN(Rendering側で既定値を補う)。
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                PresetByName(presetName),
                adjustValueList: ShapeWorkbookFixtures.AdjustValues(("adj1", 18750), ("adj2", -8333), ("adj3", 112500)));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Equal(expected, shape.Preset);
                Assert.Equal(expectedGuideCount, shape.AdjustmentValues.Count);
                Assert.Equal(0.1875, shape.AdjustmentValues[0], 4);
                Assert.Equal(-0.08333, shape.AdjustmentValues[1], 4);
                Assert.Equal(1.125, shape.AdjustmentValues[2], 4);
                Assert.True(double.IsNaN(shape.AdjustmentValues[3]));
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static A.ShapeTypeValues PresetByName(string name) => name switch
        {
            "Callout1" => A.ShapeTypeValues.Callout1,
            "Callout3" => A.ShapeTypeValues.Callout3,
            "BorderCallout1" => A.ShapeTypeValues.BorderCallout1,
            "BorderCallout2" => A.ShapeTypeValues.BorderCallout2,
            "BorderCallout3" => A.ShapeTypeValues.BorderCallout3,
            "AccentCallout2" => A.ShapeTypeValues.AccentCallout2,
            "AccentBorderCallout2" => A.ShapeTypeValues.AccentBorderCallout2,
            _ => throw new ArgumentOutOfRangeException(nameof(name), name, null),
        };

        [Fact]
        public void 雲形吹き出しのadj1とadj2が無い場合はNaNとして返す()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.CloudCallout);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Equal(2, shape.AdjustmentValues.Count);
                Assert.True(double.IsNaN(shape.AdjustmentValues[0]));
                Assert.True(double.IsNaN(shape.AdjustmentValues[1]));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void twoCellAnchorの図形を対角セルとして読み取る()
        {
            var anchor = ShapeWorkbookFixtures.TwoCellShapeAnchor(A.ShapeTypeValues.Ellipse);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shapeModel = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Equal(CellAddress.Parse("B3"), shapeModel.AnchorCell);
                var extent = Assert.IsType<CellSpanAnchorExtent>(shapeModel.Extent);
                Assert.Equal(CellAddress.Parse("D5"), extent.ToCell);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 非対応プリセットの図形はunsupportedElementsがignoreなら無視される()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.FlowChartPreparation);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects.OfType<ShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 非対応プリセットの図形はunsupportedElementsがerrorなら例外になる()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.FlowChartPreparation);
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
        public void 図形のみのシートはunsupportedElementsがerrorでも例外にならない()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<ShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 調整ガイド値を読み取り指定順に並べる()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                A.ShapeTypeValues.RoundRectangle,
                adjustValueList: ShapeWorkbookFixtures.AdjustValues(("adj", 25000)));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                var value = Assert.Single(shape.AdjustmentValues);
                Assert.Equal(0.25, value, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ガイド値が無い場合はNaNとして返す()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.RoundRectangle);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                var value = Assert.Single(shape.AdjustmentValues);
                Assert.True(double.IsNaN(value));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 回転角を度に変換して読み取る()
        {
            // 60,000分の1度単位。45度 = 2,700,000。
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, rotationEmu: 2_700_000);
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Equal(45.0, shape.RotationDegrees, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グラデーション塗りの開始色終了色と角度を読み取る()
        {
            // 60,000分の1度単位。90度 = 5,400,000。
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                A.ShapeTypeValues.Rectangle,
                fill: ShapeWorkbookFixtures.GradientFill("0000FF", "FFFFFF", 5_400_000));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                var fill = Assert.IsType<LinearGradientShapeFill>(shape.Fill);
                Assert.Equal(2, fill.Stops.Count);
                Assert.Equal(new ArgbColor(0xFF, 0x00, 0x00, 0xFF), fill.Stops[0].Color);
                Assert.Equal(new ArgbColor(0xFF, 0xFF, 0xFF, 0xFF), fill.Stops[1].Color);
                Assert.Equal(90.0, fill.AngleDegrees, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void noFillの図形は塗りつぶし無しとして読み取る()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, fill: new A.NoFill());
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.Null(shape.Fill);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図形内テキストの段落と配置を読み取る()
        {
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                A.ShapeTypeValues.Rectangle,
                textBody: ShapeWorkbookFixtures.TextBody(
                    "こんにちは",
                    vAlign: A.TextAnchoringTypeValues.Center,
                    hAlign: A.TextAlignmentTypeValues.Right,
                    fontSizeHundredthsPt: 1400,
                    bold: true));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var shape = Assert.Single(sheet.DrawingObjects.OfType<ShapeModel>());

                Assert.NotNull(shape.Text);
                Assert.Equal(VerticalAlignment.Center, shape.Text!.VAlign);
                var paragraph = Assert.Single(shape.Text.Paragraphs);
                Assert.Equal(HorizontalAlignment.Right, paragraph.HAlign);
                var run = Assert.Single(paragraph.Runs);
                Assert.Equal("こんにちは", run.Text);
                Assert.Equal(14.0, run.Font.SizePt, 3);
                Assert.True(run.Font.Bold);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図形の数が上限を超える場合はunsupportedElementsがignoreなら上限までしか読み取らない()
        {
            var count = OpenXmlWorkbookReader.MaxShapesPerSheet + 5;
            var anchors = Enumerable.Range(0, count)
                .Select(i => (DocumentFormat.OpenXml.OpenXmlElement)ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, row: i + 1, id: (uint)(i + 2)))
                .ToArray();
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchors);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(OpenXmlWorkbookReader.MaxShapesPerSheet, sheet.DrawingObjects.OfType<ShapeModel>().Count());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図形の数が上限を超える場合はunsupportedElementsがerrorなら例外になる()
        {
            var count = OpenXmlWorkbookReader.MaxShapesPerSheet + 5;
            var anchors = Enumerable.Range(0, count)
                .Select(i => (DocumentFormat.OpenXml.OpenXmlElement)ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Rectangle, row: i + 1, id: (uint)(i + 2)))
                .ToArray();
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchors);
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
        public void 図形内テキストの文字数が上限を超える場合はunsupportedElementsがignoreなら無視される()
        {
            var longText = new string('あ', 2001);
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                A.ShapeTypeValues.Rectangle,
                textBody: ShapeWorkbookFixtures.TextBody(longText));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects.OfType<ShapeModel>());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 図形内テキストの文字数が上限を超える場合はunsupportedElementsがerrorなら例外になる()
        {
            var longText = new string('あ', 2001);
            var anchor = ShapeWorkbookFixtures.ShapeAnchor(
                A.ShapeTypeValues.Rectangle,
                textBody: ShapeWorkbookFixtures.TextBody(longText));
            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("ShapeTextTooLong", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void grpSpPrにa_xfrmが無いグループは壊れたアンカーとして無視される()
        {
            // xdr:grpSpは要件10.10により構造的にサポート対象となったため「Drawing」例外は
            // もう発生しない。ただしa:xfrm(グループの子座標空間)が無ければ位置・サイズが
            // 決定できないため、ToMarker/Extent欠落時の画像・図形と同じ方針(壊れたアンカーは
            // Error/Ignoreいずれのモードでも静かに無視する)に従う。
            var groupShape = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGroupShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2U, Name = "Group" },
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGroupShapeDrawingProperties()),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShapeProperties());

            var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"),
                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent { Cx = 914400L, Cy = 914400L },
                groupShape,
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());

            var path = ShapeWorkbookFixtures.CreateWorkbook(anchor);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                Assert.Empty(workbook.Sheets[0].DrawingObjects);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像と図形が混在する場合はdrawing_xmlの出現順を保つ()
        {
            var shapeAnchor = ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.Ellipse, row: 1, id: 3U);
            var path = ImageThenShapeWorkbook(shapeAnchor);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);

                Assert.Collection(
                    sheet.DrawingObjects,
                    obj => Assert.IsType<ImageModel>(obj),
                    obj => Assert.IsType<ShapeModel>(obj));
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static string ImageThenShapeWorkbook(DocumentFormat.OpenXml.OpenXmlElement shapeAnchor)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-shape-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

            using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(
                path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                var worksheetPart = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData());

                var sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
                sheets.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "テストシート",
                });

                var drawingsPart = worksheetPart.AddNewPart<DocumentFormat.OpenXml.Packaging.DrawingsPart>();
                var imagePart = drawingsPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                using (var stream = new MemoryStream(ImageWorkbookFixtures.TinyPng()))
                {
                    imagePart.FeedData(stream);
                }

                var imageAnchor = new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                    new Xdr.Extent { Cx = 100000L, Cy = 100000L },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Image" },
                            new Xdr.NonVisualPictureDrawingProperties()),
                        new Xdr.BlipFill(
                            new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart) },
                            new A.Stretch(new A.FillRectangle())),
                        new Xdr.ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 100000L, Cy = 100000L }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                    new Xdr.ClientData());

                var drawing = new Xdr.WorksheetDrawing();
                drawing.Append(imageAnchor);
                drawing.Append(shapeAnchor);
                drawingsPart.WorksheetDrawing = drawing;

                worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            return path;
        }
    }
}
