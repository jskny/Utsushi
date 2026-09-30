using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 入力の構造に対する安全弁(要件6.7〜6.10)の結合テスト用に、最小構成の .xlsx を組み立てるヘルパー。
    /// 作成したファイルは呼び出し側で削除すること。
    /// </summary>
    internal static class SafetyLimitWorkbookFixtures
    {
        public const string SheetName = "テストシート";

        /// <summary>シート1枚(A1に値1つ)の .xlsx を作り、<paramref name="configure"/> でワークシート・ワークブックを加工する。</summary>
        public static string CreateWorkbook(Action<Workbook, Worksheet>? configure = null)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-safety-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(
                    new SheetData(
                        new Row(
                            new Cell
                            {
                                CellReference = "A1",
                                DataType = CellValues.InlineString,
                                InlineString = new InlineString(new Text("値")),
                            })
                        { RowIndex = 1U }));

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = SheetName,
                });

                configure?.Invoke(workbookPart.Workbook, worksheetPart.Worksheet);

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            return path;
        }

        /// <summary><c>&lt;col min max&gt;</c> を指定した範囲の数だけ並べた .xlsx を作る(要件6.8)。</summary>
        public static string CreateWithColumns(IEnumerable<(uint Min, uint Max)> columns) =>
            CreateWorkbook((_, worksheet) =>
            {
                var cols = new Columns();
                foreach (var (min, max) in columns)
                {
                    cols.Append(new Column { Min = min, Max = max, Width = 12.0, CustomWidth = true });
                }

                // スキーマ上 <cols> は <sheetData> の前に置く。
                worksheet.InsertBefore(cols, worksheet.GetFirstChild<SheetData>());
            });

        /// <summary>手動改ページ(<c>&lt;brk man="1"&gt;</c>)を行・列それぞれ指定数だけ持つ .xlsx を作る(要件6.8)。</summary>
        public static string CreateWithBreaks(int rowBreakCount, int columnBreakCount) =>
            CreateWorkbook((_, worksheet) =>
            {
                if (rowBreakCount > 0)
                {
                    var rowBreaks = new RowBreaks { Count = (uint)rowBreakCount, ManualBreakCount = (uint)rowBreakCount };
                    for (var i = 1; i <= rowBreakCount; i++)
                    {
                        rowBreaks.Append(new Break { Id = (uint)i, Max = 16383U, ManualPageBreak = true });
                    }

                    worksheet.Append(rowBreaks);
                }

                if (columnBreakCount > 0)
                {
                    var columnBreaks = new ColumnBreaks { Count = (uint)columnBreakCount, ManualBreakCount = (uint)columnBreakCount };
                    for (var i = 1; i <= columnBreakCount; i++)
                    {
                        columnBreaks.Append(new Break { Id = (uint)i, Max = 1048575U, ManualPageBreak = true });
                    }

                    worksheet.Append(columnBreaks);
                }
            });

        /// <summary>1セルずつの印刷範囲(<c>_xlnm.Print_Area</c>)を指定個数だけカンマ区切りで持つ .xlsx を作る(要件6.8)。</summary>
        public static string CreateWithPrintAreas(int count) =>
            CreateWorkbook((workbook, _) =>
            {
                var references = string.Join(
                    ",",
                    Enumerable.Range(1, count).Select(i =>
                        "'" + SheetName + "'!$A$" + i.ToString(CultureInfo.InvariantCulture)
                        + ":$A$" + i.ToString(CultureInfo.InvariantCulture)));

                workbook.Append(new DefinedNames(
                    new DefinedName(references) { Name = "_xlnm.Print_Area", LocalSheetId = 0U }));
            });

        /// <summary>
        /// 参照先パートが存在しない <c>r:embed</c> を持つ画像(<c>xdr:pic</c>)の oneCellAnchor を組み立てる(要件6.10)。
        /// <see cref="ShapeWorkbookFixtures.CreateWorkbook"/> と組み合わせる(図形用の drawing パートには画像パートが無い)。
        /// </summary>
        public static Xdr.OneCellAnchor PictureAnchorWithMissingEmbed(string embedId = "rIdMissing", uint id = 3U)
        {
            const long widthEmu = 60L * 12700L;
            const long heightEmu = 20L * 12700L;

            return new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = id, Name = "MissingImage" },
                        new Xdr.NonVisualPictureDrawingProperties()),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = embedId },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData());
        }

        /// <summary>.xlsx(ZIP)内の指定エントリの中身を、任意の文字列(壊れたXML等)に差し替える。</summary>
        public static void ReplaceEntry(string path, string entryName, string content)
        {
            using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
            var entry = archive.GetEntry(entryName)
                ?? throw new InvalidOperationException($"エントリ '{entryName}' がありません。");
            entry.Delete();

            var replaced = archive.CreateEntry(entryName);
            using var stream = replaced.Open();
            var bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(content);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>.xlsx(ZIP)内の指定エントリの中身(UTF-8のテキスト)を <paramref name="transform"/> で書き換える。</summary>
        public static void ModifyEntry(string path, string entryName, Func<string, string> transform)
        {
            string original;
            using (var archive = ZipFile.OpenRead(path))
            {
                var entry = archive.GetEntry(entryName)
                    ?? throw new InvalidOperationException($"エントリ '{entryName}' がありません。");
                using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
                original = reader.ReadToEnd();
            }

            var modified = transform(original);
            if (string.Equals(original, modified, StringComparison.Ordinal))
            {
                throw new InvalidOperationException($"エントリ '{entryName}' の書き換えで内容が変わりませんでした。");
            }

            ReplaceEntry(path, entryName, modified);
        }

        /// <summary>
        /// .xlsx(ZIP)内の指定エントリの圧縮データを、中央ディレクトリ(宣言サイズ)はそのままに 0xFF で塗りつぶす。
        /// deflate の先頭ブロックが予約済みのブロック種別(BTYPE=11)になり、展開時に
        /// <see cref="InvalidDataException"/> になる(要件6.7。壊れた圧縮データの再現)。
        /// </summary>
        public static void CorruptEntryCompressedData(string path, string entryName)
        {
            long compressedLength;
            using (var archive = ZipFile.OpenRead(path))
            {
                var entry = archive.GetEntry(entryName)
                    ?? throw new InvalidOperationException($"エントリ '{entryName}' がありません。");
                compressedLength = entry.CompressedLength;
                if (compressedLength == entry.Length)
                {
                    throw new InvalidOperationException($"エントリ '{entryName}' が圧縮されていません。");
                }
            }

            var bytes = File.ReadAllBytes(path);
            var name = Encoding.UTF8.GetBytes(entryName);
            for (var i = 0; i + 30 <= bytes.Length; i++)
            {
                // ローカルファイルヘッダー(PK\x03\x04)。ファイル名の長さは26バイト目、拡張フィールドの長さは28バイト目。
                if (bytes[i] != 0x50 || bytes[i + 1] != 0x4B || bytes[i + 2] != 0x03 || bytes[i + 3] != 0x04)
                {
                    continue;
                }

                var nameLength = bytes[i + 26] | (bytes[i + 27] << 8);
                var extraLength = bytes[i + 28] | (bytes[i + 29] << 8);
                if (nameLength != name.Length || !bytes.AsSpan(i + 30, nameLength).SequenceEqual(name))
                {
                    continue;
                }

                var dataStart = i + 30 + nameLength + extraLength;
                for (var j = 0L; j < compressedLength; j++)
                {
                    bytes[dataStart + j] = 0xFF;
                }

                File.WriteAllBytes(path, bytes);
                return;
            }

            throw new InvalidOperationException($"エントリ '{entryName}' のローカルファイルヘッダーが見つかりません。");
        }

        /// <summary>A1 のインライン文字列を <paramref name="text"/> にした .xlsx を作る(要件6.8)。</summary>
        public static string CreateWithInlineText(string text) =>
            CreateWorkbook((_, worksheet) =>
            {
                var cell = worksheet.Descendants<Cell>().First();
                cell.InlineString = new InlineString(new Text(text));
            });

        /// <summary>
        /// 共有文字列 <paramref name="items"/> を持ち、A2 が先頭の共有文字列を参照する .xlsx を作る(要件6.8)。
        /// </summary>
        public static string CreateWithSharedStrings(params SharedStringItem[] items) =>
            CreateWorkbook((workbook, worksheet) =>
            {
                var sharedStringPart = workbook.WorkbookPart!.AddNewPart<SharedStringTablePart>();
                sharedStringPart.SharedStringTable = new SharedStringTable(items);
                sharedStringPart.SharedStringTable.Save();

                worksheet.GetFirstChild<SheetData>()!.Append(
                    new Row(new Cell { CellReference = "A2", DataType = CellValues.SharedString, CellValue = new CellValue("0") })
                    {
                        RowIndex = 2U,
                    });
            });

        /// <summary><c>&lt;headerFooter&gt;</c> に <paramref name="element"/>(<c>oddHeader</c> 等)を1つだけ持つ .xlsx を作る(要件6.8)。</summary>
        public static string CreateWithHeaderFooter(OpenXmlLeafTextElement element) =>
            CreateWorkbook((_, worksheet) => worksheet.Append(new HeaderFooter(element)));

        /// <summary><c>pageSetup/@scale</c> を指定した .xlsx を作る(要件6.9)。</summary>
        public static string CreateWithScale(uint scale) =>
            CreateWorkbook((_, worksheet) => worksheet.Append(new PageSetup { Scale = scale }));

        /// <summary>
        /// ユーザー定義の数値書式(<c>numFmtId=164</c>)を <paramref name="formatCode"/> とし、A2 の数値 1234.5 に
        /// 適用した .xlsx を作る(要件6.8)。
        /// </summary>
        public static string CreateWithNumberFormat(string formatCode) =>
            CreateWorkbook((workbook, worksheet) =>
            {
                var stylesPart = workbook.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(
                    new NumberingFormats(new NumberingFormat { NumberFormatId = 164U, FormatCode = formatCode }) { Count = 1U },
                    new Fonts(new Font()) { Count = 1U },
                    new Fills(new Fill()) { Count = 1U },
                    new Borders(new Border()) { Count = 1U },
                    new CellFormats(
                        new CellFormat(),
                        new CellFormat { NumberFormatId = 164U, ApplyNumberFormat = true })
                    { Count = 2U });
                stylesPart.Stylesheet.Save();

                worksheet.GetFirstChild<SheetData>()!.Append(
                    new Row(new Cell { CellReference = "A2", StyleIndex = 1U, CellValue = new CellValue("1234.5") })
                    {
                        RowIndex = 2U,
                    });
            });

        /// <summary>
        /// ワークブックの関係パート(<c>xl/_rels/workbook.xml.rels</c>)に、外部ハイパーリンクの関係を
        /// <paramref name="count"/> 件追加する(要件6.7)。既存の関係(ワークシート1件)はそのまま残す。
        /// </summary>
        public static void AddExternalRelationships(string path, string relsEntryName, int count) =>
            ModifyEntry(path, relsEntryName, xml =>
            {
                var builder = new StringBuilder(count * 160);
                for (var i = 0; i < count; i++)
                {
                    var n = i.ToString(CultureInfo.InvariantCulture);
                    builder.Append("<Relationship Id=\"rIdExt").Append(n)
                        .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\"")
                        .Append(" Target=\"https://example.invalid/").Append(n).Append("\" TargetMode=\"External\"/>");
                }

                var end = xml.LastIndexOf("</Relationships>", StringComparison.Ordinal);
                return xml.Substring(0, end) + builder + xml.Substring(end);
            });

        /// <summary>.xlsx(ZIP)内のエントリの中から、名前が <paramref name="suffix"/> で終わる最初のものを返す。</summary>
        public static string EntryNameEndingWith(string path, string suffix)
        {
            using var archive = ZipFile.OpenRead(path);
            return archive.Entries
                .Select(e => e.FullName)
                .First(n => n.EndsWith(suffix, StringComparison.Ordinal));
        }

        /// <summary>最初のワークシートパートのZIPエントリ名を返す。</summary>
        public static string FirstWorksheetEntryName(string path)
        {
            using var archive = ZipFile.OpenRead(path);
            return archive.Entries
                .Select(e => e.FullName)
                .First(n => n.StartsWith("xl/worksheets/", StringComparison.Ordinal) && n.EndsWith(".xml", StringComparison.Ordinal));
        }
    }
}
