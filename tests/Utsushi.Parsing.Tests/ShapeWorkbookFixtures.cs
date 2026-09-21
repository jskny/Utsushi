using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 図形(要件10)の読み取りテスト用に、最小構成の .xlsx を組み立てるヘルパー。
    /// </summary>
    internal static class ShapeWorkbookFixtures
    {
        /// <summary>指定したアンカー群を含む最小の .xlsx を一時ファイルとして作成する。呼び出し側で削除すること。</summary>
        public static string CreateWorkbook(params OpenXmlElement[] anchors)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-shape-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "テストシート",
                });

                var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                var drawing = new Xdr.WorksheetDrawing();
                foreach (var anchor in anchors)
                {
                    drawing.Append(anchor);
                }

                drawingsPart.WorksheetDrawing = drawing;

                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            return path;
        }

        /// <summary>指定した設定で <c>xdr:sp</c> の oneCellAnchor を組み立てる。</summary>
        public static Xdr.OneCellAnchor ShapeAnchor(
            A.ShapeTypeValues preset,
            int row = 3,
            int column = 2,
            long widthEmu = 900000L,
            long heightEmu = 900000L,
            int rotationEmu = 0,
            OpenXmlElement? fill = null,
            A.Outline? outline = null,
            A.AdjustValueList? adjustValueList = null,
            Xdr.TextBody? textBody = null,
            uint id = 2U)
        {
            var spPrChildren = new List<OpenXmlElement>
            {
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = widthEmu, Cy = heightEmu })
                {
                    Rotation = rotationEmu,
                },
                new A.PresetGeometry(adjustValueList ?? new A.AdjustValueList()) { Preset = preset },
            };

            if (fill is not null)
            {
                spPrChildren.Add(fill);
            }

            if (outline is not null)
            {
                spPrChildren.Add(outline);
            }

            var shapeChildren = new List<OpenXmlElement>
            {
                new Xdr.NonVisualShapeProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "Shape" + id.ToString(CultureInfo.InvariantCulture) },
                    new Xdr.NonVisualShapeDrawingProperties()),
                new Xdr.ShapeProperties(spPrChildren),
            };

            if (textBody is not null)
            {
                shapeChildren.Add(textBody);
            }

            return new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
                new Xdr.Shape(shapeChildren),
                new Xdr.ClientData());
        }

        /// <summary>指定した対角セルまでの <c>xdr:sp</c> の twoCellAnchor を組み立てる。</summary>
        public static Xdr.TwoCellAnchor TwoCellShapeAnchor(
            A.ShapeTypeValues preset,
            int fromRow = 3,
            int fromColumn = 2,
            int toRow = 5,
            int toColumn = 4,
            uint id = 2U)
        {
            var shapeChildren = new List<OpenXmlElement>
            {
                new Xdr.NonVisualShapeProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "Shape" + id.ToString(CultureInfo.InvariantCulture) },
                    new Xdr.NonVisualShapeDrawingProperties()),
                new Xdr.ShapeProperties(new A.PresetGeometry(new A.AdjustValueList()) { Preset = preset }),
            };

            return new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId((fromColumn - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((fromRow - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId((toColumn - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((toRow - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(shapeChildren),
                new Xdr.ClientData());
        }

        /// <summary>単色塗りつぶし(<c>a:solidFill</c>)。</summary>
        public static A.SolidFill SolidFill(string hex) => new(new A.RgbColorModelHex { Val = hex });

        /// <summary>線形グラデーション塗りつぶし(<c>a:gradFill</c>)。</summary>
        public static A.GradientFill GradientFill(string startHex, string endHex, int angleEmu) => new(
            new A.GradientStopList(
                new A.GradientStop(new A.RgbColorModelHex { Val = startHex }) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = endHex }) { Position = 100000 }),
            new A.LinearGradientFill { Angle = angleEmu });

        /// <summary>枠線(<c>a:ln</c>)。</summary>
        public static A.Outline Outline(string hex, long widthEmu) =>
            new(new A.SolidFill(new A.RgbColorModelHex { Val = hex })) { Width = (int)widthEmu };

        /// <summary>
        /// 単一段落・単一ランのテキスト(<c>xdr:txBody</c>)を組み立てる。
        /// </summary>
        public static Xdr.TextBody TextBody(
            string text,
            A.TextAnchoringTypeValues vAlign = A.TextAnchoringTypeValues.Top,
            A.TextAlignmentTypeValues hAlign = A.TextAlignmentTypeValues.Left,
            int fontSizeHundredthsPt = 1000,
            bool bold = false) =>
            new(
                new A.BodyProperties { Anchor = vAlign },
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Alignment = hAlign },
                    new A.Run(
                        new A.RunProperties { FontSize = fontSizeHundredthsPt, Bold = bold },
                        new A.Text(text))));

        /// <summary>調整ガイド1つ(<c>a:gd</c>)を持つ <c>a:avLst</c>。</summary>
        public static A.AdjustValueList AdjustValues(params (string Name, int Value)[] guides)
        {
            var list = new A.AdjustValueList();
            foreach (var (name, value) in guides)
            {
                list.Append(new A.ShapeGuide { Name = name, Formula = "val " + value.ToString(CultureInfo.InvariantCulture) });
            }

            return list;
        }
    }
}
