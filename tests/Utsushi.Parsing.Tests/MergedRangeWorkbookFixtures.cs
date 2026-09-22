using System;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 結合セル範囲(要件2.9)の読み取りテスト用に、最小構成の .xlsx を組み立てるヘルパー。
    /// </summary>
    internal static class MergedRangeWorkbookFixtures
    {
        /// <summary>
        /// <paramref name="count"/>個の重複しない2セル結合範囲(<c>A{i}:B{i}</c>。1行に1つ)を
        /// 含む .xlsx を一時ファイルとして作成し、そのパスを返す。呼び出し側で削除すること。
        /// 結合範囲の数の上限(<see cref="OpenXmlWorkbookReader.MaxMergedRangesPerSheet"/>)のテスト用。
        /// </summary>
        public static string CreateWithManyMergedRanges(int count)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-merged-range-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                var mergeCells = new MergeCells();

                for (var i = 1; i <= count; i++)
                {
                    var row = i.ToString(CultureInfo.InvariantCulture);
                    mergeCells.Append(new MergeCell { Reference = $"A{row}:B{row}" });
                }

                worksheetPart.Worksheet = new Worksheet(sheetData, mergeCells);

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "テストシート",
                });

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            return path;
        }
    }
}
