using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// アクティブシートの解決(要件12.2)のテスト用に、複数シート・<c>workbookView/@activeTab</c>・
    /// シートの表示状態(<c>sheet/@state</c>)を指定できる最小構成の .xlsx を組み立てるヘルパー。
    /// </summary>
    internal static class ActiveSheetWorkbookFixtures
    {
        /// <summary>1枚のシートの指定。</summary>
        /// <param name="Name">シート名。A1セルにもこの名前を文字列として書き込む。</param>
        /// <param name="State">シートの表示状態。null なら属性を省略する(=表示)。</param>
        /// <param name="Anchors">シートに置く描画オブジェクトのアンカー。null または空なら描画パートを作らない。</param>
        public sealed record SheetSpec(
            string Name,
            SheetStateValues? State = null,
            IReadOnlyList<OpenXmlElement>? Anchors = null);

        /// <summary>
        /// 指定したシート群を持つ .xlsx を一時ファイルとして作成する。呼び出し側で削除すること。
        /// </summary>
        /// <param name="activeTab">
        /// <c>workbookView/@activeTab</c> の値。null なら <c>bookViews</c> 要素自体を省略する。
        /// </param>
        public static string CreateWorkbook(uint? activeTab, params SheetSpec[] sheetSpecs)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-active-sheet-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // スキーマ上 bookViews は sheets より前に置く。
                if (activeTab is { } tab)
                {
                    workbookPart.Workbook.Append(new BookViews(new WorkbookView { ActiveTab = tab }));
                }

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                var sheetId = 1U;
                foreach (var spec in sheetSpecs)
                {
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(
                        new SheetData(
                            new Row(
                                new Cell(new InlineString(new Text(spec.Name)))
                                {
                                    CellReference = "A1",
                                    DataType = CellValues.InlineString,
                                })
                            { RowIndex = 1U }));

                    if (spec.Anchors is { Count: > 0 } anchors)
                    {
                        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                        var drawing = new Xdr.WorksheetDrawing();
                        foreach (var anchor in anchors)
                        {
                            drawing.Append(anchor);
                        }

                        drawingsPart.WorksheetDrawing = drawing;
                        worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                    }

                    worksheetPart.Worksheet.Save();

                    var sheet = new Sheet
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId++,
                        Name = spec.Name,
                    };

                    if (spec.State is { } state)
                    {
                        sheet.State = state;
                    }

                    sheets.Append(sheet);
                }

                workbookPart.Workbook.Save();
            }

            return path;
        }
    }
}
