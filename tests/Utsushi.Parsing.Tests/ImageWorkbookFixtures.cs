using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 画像(要件9)の読み取りテスト用に、最小構成の .xlsx を組み立てるヘルパー。
    /// </summary>
    internal static class ImageWorkbookFixtures
    {
        private const double EmusPerPoint = 12700.0;

        /// <summary>
        /// 画像1枚(<paramref name="useTwoCellAnchor"/>に応じて oneCellAnchor / twoCellAnchor)を
        /// 含む最小の .xlsx を一時ファイルとして作成し、そのパスを返す。呼び出し側で削除すること。
        /// </summary>
        public static string CreateWithPicture(
            string contentType, byte[] imageData, bool useTwoCellAnchor = false, ImagePartType partType = ImagePartType.Png)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-image-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

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

                // ImagePartは実際のContentTypeを指定できないケース(未対応形式のテスト)があるため、
                // FeedDataでバイト列だけを流し込み、ContentTypeはPngパートの型を借りて上書きする。
                var imagePart = drawingsPart.AddImagePart(partType);
                using (var stream = new MemoryStream(imageData))
                {
                    imagePart.FeedData(stream);
                }

                var drawing = new Xdr.WorksheetDrawing();
                drawing.Append(useTwoCellAnchor ? BuildTwoCellAnchor(drawingsPart, imagePart) : BuildOneCellAnchor(drawingsPart, imagePart));
                drawingsPart.WorksheetDrawing = drawing;

                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            OverrideContentType(path, contentType);
            return path;
        }

        /// <summary>画像を含まない図形(<c>xdr:sp</c>)だけを含む最小の .xlsx を作る。</summary>
        public static string CreateWithNonPictureShape()
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-image-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

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

                var shape = new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Rectangle" },
                        new Xdr.NonVisualShapeDrawingProperties()),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 914400L, Cy = 914400L }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

                var anchor = new Xdr.OneCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                    new Xdr.Extent { Cx = 914400L, Cy = 914400L },
                    shape,
                    new Xdr.ClientData());

                var drawing = new Xdr.WorksheetDrawing();
                drawing.Append(anchor);
                drawingsPart.WorksheetDrawing = drawing;

                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            return path;
        }

        private static Xdr.OneCellAnchor BuildOneCellAnchor(DrawingsPart drawingsPart, ImagePart imagePart)
        {
            const long widthEmu = 60L * (long)EmusPerPoint;
            const long heightEmu = 20L * (long)EmusPerPoint;

            return new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"), // 0始まり → CellAddress上は列C
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"), // 0始まり → CellAddress上は行3
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
                BuildPicture(drawingsPart, imagePart, widthEmu, heightEmu),
                new Xdr.ClientData());
        }

        private static Xdr.TwoCellAnchor BuildTwoCellAnchor(DrawingsPart drawingsPart, ImagePart imagePart)
        {
            return new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"), new Xdr.ColumnOffset("0"), new Xdr.RowId("2"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("3"), // 0始まり → 列E
                    new Xdr.ColumnOffset((10L * (long)EmusPerPoint).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowId("4"), // 0始まり → 行5
                    new Xdr.RowOffset((5L * (long)EmusPerPoint).ToString(CultureInfo.InvariantCulture))),
                BuildPicture(drawingsPart, imagePart, 0L, 0L),
                new Xdr.ClientData());
        }

        private static Xdr.Picture BuildPicture(DrawingsPart drawingsPart, ImagePart imagePart, long widthEmu, long heightEmu) =>
            new(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Logo" },
                    new Xdr.NonVisualPictureDrawingProperties()),
                new Xdr.BlipFill(
                    new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart) },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        /// <summary>
        /// 未対応形式のテスト用に、ImagePartのContentTypeを実際のバイト列とは無関係な値に差し替える。
        /// Open XML SDKはパート作成時に <see cref="ImagePartType"/> からContentTypeを決めるため、
        /// パッケージ内のXMLを直接書き換える。
        /// </summary>
        private static void OverrideContentType(string path, string contentType)
        {
            if (string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            using var document = SpreadsheetDocument.Open(path, isEditable: true);
            var drawingsPart = document.WorkbookPart!.WorksheetParts.First().DrawingsPart!;
            var imagePart = drawingsPart.ImageParts.First();

            // ImagePart.ContentTypeは読み取り専用のため、既存パートを削除して差し替える。
            // DeletePartの前にストリームを確実に閉じておく必要がある(開いたままだと削除に失敗する)。
            using var buffer = new MemoryStream();
            using (var data = imagePart.GetStream())
            {
                data.CopyTo(buffer);
            }

            buffer.Position = 0;

            var relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.DeletePart(imagePart);

            var newPart = drawingsPart.AddExtendedPart(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                contentType,
                GetExtension(contentType),
                relationshipId);
            newPart.FeedData(buffer);
        }

        private static string GetExtension(string contentType) => contentType switch
        {
            "image/tiff" => ".tiff",
            "image/x-emf" => ".emf",
            _ => ".bin",
        };
    }
}
