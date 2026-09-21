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
    /// 接続線(<c>xdr:cxnSp</c>。要件10.9)とグループ化された図形(<c>xdr:grpSp</c>。要件10.10)の
    /// 読み取りテスト用に、最小構成の要素/.xlsx を組み立てるヘルパー。
    /// </summary>
    internal static class ConnectorAndGroupWorkbookFixtures
    {
        /// <summary>
        /// <paramref name="buildTopLevelAnchor"/> が返すアンカー1つだけを含む最小の .xlsx を
        /// 一時ファイルとして作成する。グループ内に画像を1枚含める場合など、アンカー構築時に
        /// <see cref="DrawingsPart"/>/<see cref="ImagePart"/>(relationship id解決)が必要なテスト用。
        /// 呼び出し側でファイルを削除すること。
        /// </summary>
        public static string CreateWorkbookWithImage(
            System.Func<DrawingsPart, ImagePart, OpenXmlElement> buildTopLevelAnchor)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-conn-group-test-" + System.Guid.NewGuid().ToString("N") + ".xlsx");

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
                var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
                using (var stream = new MemoryStream(ImageWorkbookFixtures.TinyPng()))
                {
                    imagePart.FeedData(stream);
                }

                var anchor = buildTopLevelAnchor(drawingsPart, imagePart);
                var drawing = new Xdr.WorksheetDrawing();
                drawing.Append(anchor);
                drawingsPart.WorksheetDrawing = drawing;

                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            return path;
        }

        /// <summary>指定した設定で <c>xdr:cxnSp</c> の oneCellAnchor を組み立てる(要件10.9)。</summary>
        public static Xdr.OneCellAnchor ConnectorAnchor(
            A.ShapeTypeValues preset,
            int row = 3,
            int column = 2,
            long widthEmu = 900000L,
            long heightEmu = 900000L,
            int rotationEmu = 0,
            bool flipH = false,
            bool flipV = false,
            A.Outline? outline = null,
            uint id = 2U)
        {
            var spPrChildren = new List<OpenXmlElement>
            {
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = widthEmu, Cy = heightEmu })
                {
                    Rotation = rotationEmu,
                    HorizontalFlip = flipH,
                    VerticalFlip = flipV,
                },
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = preset },
            };

            if (outline is not null)
            {
                spPrChildren.Add(outline);
            }

            var connector = new Xdr.ConnectionShape(
                new Xdr.NonVisualConnectionShapeProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "Connector" + id.ToString(CultureInfo.InvariantCulture) },
                    new Xdr.NonVisualConnectorShapeDrawingProperties()),
                new Xdr.ShapeProperties(spPrChildren));

            return WrapInOneCellAnchor(connector, row, column, widthEmu, heightEmu);
        }

        /// <summary>グループ(<c>xdr:grpSp</c>)本体を組み立てる。子座標空間(<c>a:chOff</c>/<c>a:chExt</c>)のみを持ち、
        /// グループ自身の位置はセルアンカー側で決まる(トップレベルのグループ用。要件10.10)。</summary>
        public static Xdr.GroupShape GroupShapeElement(
            IReadOnlyList<OpenXmlElement> children,
            long chOffX = 0L,
            long chOffY = 0L,
            long chExtCx = 1800000L,
            long chExtCy = 1800000L,
            int rotationEmu = 0,
            uint id = 2U)
        {
            var transformGroup = new A.TransformGroup(
                new A.ChildOffset { X = chOffX, Y = chOffY },
                new A.ChildExtents { Cx = chExtCx, Cy = chExtCy })
            {
                Rotation = rotationEmu,
            };

            var groupChildren = new List<OpenXmlElement>
            {
                new Xdr.NonVisualGroupShapeProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "Group" + id.ToString(CultureInfo.InvariantCulture) },
                    new Xdr.NonVisualGroupShapeDrawingProperties()),
                new Xdr.GroupShapeProperties(transformGroup),
            };
            groupChildren.AddRange(children);

            return new Xdr.GroupShape(groupChildren.ToArray());
        }

        /// <summary>入れ子の<c>xdr:grpSp</c>(グループ内グループ)を組み立てる。親の子座標空間上での
        /// 自身の位置・サイズ(<c>a:off</c>/<c>a:ext</c>)と、自身の子孫のための子座標空間
        /// (<c>a:chOff</c>/<c>a:chExt</c>)の両方を持つ(要件10.10)。</summary>
        public static Xdr.GroupShape NestedGroupElement(
            IReadOnlyList<OpenXmlElement> children,
            long offX,
            long offY,
            long extCx,
            long extCy,
            long chOffX = 0L,
            long chOffY = 0L,
            long chExtCx = 900000L,
            long chExtCy = 900000L,
            int rotationEmu = 0,
            uint id = 30U)
        {
            var transformGroup = new A.TransformGroup(
                new A.Offset { X = offX, Y = offY },
                new A.Extents { Cx = extCx, Cy = extCy },
                new A.ChildOffset { X = chOffX, Y = chOffY },
                new A.ChildExtents { Cx = chExtCx, Cy = chExtCy })
            {
                Rotation = rotationEmu,
            };

            var groupChildren = new List<OpenXmlElement>
            {
                new Xdr.NonVisualGroupShapeProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "NestedGroup" + id.ToString(CultureInfo.InvariantCulture) },
                    new Xdr.NonVisualGroupShapeDrawingProperties()),
                new Xdr.GroupShapeProperties(transformGroup),
            };
            groupChildren.AddRange(children);

            return new Xdr.GroupShape(groupChildren.ToArray());
        }

        /// <summary>
        /// <paramref name="leaf"/> を <paramref name="nestedLevels"/> 段のネストした<c>xdr:grpSp</c>で
        /// 包む。戻り値はもっとも外側(トップレベルグループの直接の子)のグループ。
        /// ネスト段数上限(<see cref="Utsushi.Parsing.OpenXml.OpenXmlWorkbookReader.MaxShapeNestingDepth"/>)の
        /// 境界値テスト用。
        /// </summary>
        public static OpenXmlElement WrapNestedGroups(int nestedLevels, OpenXmlElement leaf)
        {
            OpenXmlElement current = leaf;
            for (var i = 0; i < nestedLevels; i++)
            {
                current = NestedGroupElement(
                    new[] { current },
                    offX: 0L,
                    offY: 0L,
                    extCx: 900000L,
                    extCy: 900000L,
                    chOffX: 0L,
                    chOffY: 0L,
                    chExtCx: 900000L,
                    chExtCy: 900000L,
                    id: (uint)(500 + i));
            }

            return current;
        }

        /// <summary>グループ内の<c>xdr:sp</c>子要素を組み立てる(要件10.10)。位置は親グループの
        /// 子座標空間上の<c>a:off</c>/<c>a:ext</c>(EMU)で指定する。</summary>
        public static Xdr.Shape GroupChildShapeElement(
            A.ShapeTypeValues preset,
            long offX,
            long offY,
            long extCx,
            long extCy,
            int rotationEmu = 0,
            OpenXmlElement? fill = null,
            A.Outline? outline = null,
            Xdr.TextBody? textBody = null,
            uint id = 10U)
        {
            var spPrChildren = new List<OpenXmlElement>
            {
                new A.Transform2D(
                    new A.Offset { X = offX, Y = offY },
                    new A.Extents { Cx = extCx, Cy = extCy })
                {
                    Rotation = rotationEmu,
                },
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = preset },
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

            return new Xdr.Shape(shapeChildren.ToArray());
        }

        /// <summary>グループ内の<c>xdr:cxnSp</c>子要素を組み立てる(要件10.9, 10.10)。</summary>
        public static Xdr.ConnectionShape GroupChildConnectorElement(
            A.ShapeTypeValues preset,
            long offX,
            long offY,
            long extCx,
            long extCy,
            int rotationEmu = 0,
            bool flipH = false,
            bool flipV = false,
            A.Outline? outline = null,
            uint id = 12U)
        {
            var spPrChildren = new List<OpenXmlElement>
            {
                new A.Transform2D(
                    new A.Offset { X = offX, Y = offY },
                    new A.Extents { Cx = extCx, Cy = extCy })
                {
                    Rotation = rotationEmu,
                    HorizontalFlip = flipH,
                    VerticalFlip = flipV,
                },
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = preset },
            };

            if (outline is not null)
            {
                spPrChildren.Add(outline);
            }

            return new Xdr.ConnectionShape(
                new Xdr.NonVisualConnectionShapeProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "GroupConnector" + id.ToString(CultureInfo.InvariantCulture) },
                    new Xdr.NonVisualConnectorShapeDrawingProperties()),
                new Xdr.ShapeProperties(spPrChildren));
        }

        /// <summary>グループ内の<c>xdr:pic</c>子要素を組み立てる(要件10.10)。呼び出し側で
        /// <see cref="CreateWorkbookWithImage"/>を使って<see cref="DrawingsPart"/>/<see cref="ImagePart"/>を用意すること。</summary>
        public static Xdr.Picture GroupChildImageElement(
            DrawingsPart drawingsPart,
            ImagePart imagePart,
            long offX,
            long offY,
            long extCx,
            long extCy,
            uint id = 20U) =>
            new(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "GroupImage" + id.ToString(CultureInfo.InvariantCulture) },
                    new Xdr.NonVisualPictureDrawingProperties()),
                new Xdr.BlipFill(
                    new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart) },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = offX, Y = offY },
                        new A.Extents { Cx = extCx, Cy = extCy }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        /// <summary>
        /// グループ内の非対応の描画オブジェクト種別(<c>xdr:sp</c>/<c>xdr:pic</c>/<c>xdr:cxnSp</c>/
        /// <c>xdr:grpSp</c>のいずれでもない要素)を模した最小の<c>xdr:graphicFrame</c>を組み立てる
        /// (要件10.7)。内容の妥当性は問わない(型さえ一致すれば読み取り側は非対応として扱う)。
        /// </summary>
        public static Xdr.GraphicFrame GraphicFrameElement(uint id = 99U) => new();

        /// <summary>指定した<c>xdr:sp</c>/<c>xdr:pic</c>/<c>xdr:cxnSp</c>/<c>xdr:grpSp</c>を
        /// oneCellAnchorで包む。</summary>
        public static Xdr.OneCellAnchor WrapInOneCellAnchor(
            OpenXmlElement drawingElement,
            int row = 3,
            int column = 2,
            long widthEmu = 1800000L,
            long heightEmu = 1800000L) =>
            new(
                new Xdr.FromMarker(
                    new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
                drawingElement,
                new Xdr.ClientData());

        /// <summary>指定した<c>xdr:sp</c>/<c>xdr:pic</c>/<c>xdr:cxnSp</c>/<c>xdr:grpSp</c>を
        /// twoCellAnchorで包む(<c>HasUnsupportedDrawingObject</c>のoneCellAnchor/twoCellAnchor
        /// 両対応を確認するテスト用)。</summary>
        public static Xdr.TwoCellAnchor WrapInTwoCellAnchor(
            OpenXmlElement drawingElement,
            int fromRow = 3,
            int fromColumn = 2,
            int toRow = 5,
            int toColumn = 4) =>
            new(
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
                drawingElement,
                new Xdr.ClientData());
    }
}
