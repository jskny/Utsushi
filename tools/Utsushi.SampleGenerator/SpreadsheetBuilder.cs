using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.SampleGenerator
{
/// <summary>
/// 帳票サンプル用の .xlsx を組み立てる薄いビルダー。
/// </summary>
/// <remarks>
/// サンプル生成専用であり、Utsushi 本体のコードからは参照されない。
/// Open XML SDK の低レベルAPIをそのまま使うと記述量が多くなるため、
/// サンプル作成に必要な範囲だけをまとめている。
/// </remarks>
internal sealed class SpreadsheetBuilder
{
    /// <summary>1ポイントあたりのEMU(English Metric Unit)数。OOXML描画要素の座標・サイズの単位。</summary>
    private const double EmusPerPoint = 12700.0;

    private readonly Dictionary<string, (int Row, int Column, object? Value, uint StyleIndex, bool IsNumber)> _cells = new();
    private readonly List<string> _mergedRanges = new();
    private readonly Dictionary<int, double> _columnWidths = new();
    private readonly Dictionary<int, double> _rowHeights = new();
    private readonly List<uint> _manualRowBreaks = new();
    private (int Row, int Column, double OffsetXPt, double OffsetYPt, double WidthPt, double HeightPt, byte[] Png)? _image;

    public SpreadsheetBuilder(string sheetName)
    {
        SheetName = sheetName;
    }

    public string SheetName { get; }

    /// <summary>印刷範囲。複数指定する場合はカンマ区切り(例: "$A$1:$F$20,$A$22:$F$41")。</summary>
    public string? PrintArea { get; set; }

    public string? PrintTitleRows { get; set; }

    /// <summary>ヘッダーの書式コード(例: "&amp;L&amp;D&amp;R&amp;A")。</summary>
    public string? OddHeader { get; set; }

    /// <summary>フッターの書式コード(例: "&amp;C&amp;P / &amp;N")。</summary>
    public string? OddFooter { get; set; }

    /// <summary>先頭ページだけ別のヘッダー/フッターを使うかどうか。</summary>
    public bool DifferentFirstPage { get; set; }

    public string? FirstHeader { get; set; }

    public string? FirstFooter { get; set; }

    public uint PaperSizeCode { get; set; } = 9; // A4

    public OrientationValues Orientation { get; set; } = OrientationValues.Portrait;

    public double MarginLeftIn { get; set; } = 0.7;

    public double MarginRightIn { get; set; } = 0.7;

    public double MarginTopIn { get; set; } = 0.75;

    public double MarginBottomIn { get; set; } = 0.75;

    public uint ScalePercent { get; set; } = 100;

    public void SetColumnWidth(int column, double width) => _columnWidths[column] = width;

    public void SetRowHeight(int row, double heightPt) => _rowHeights[row] = heightPt;

    public void Merge(string range) => _mergedRanges.Add(range);

    public void AddManualRowBreak(uint rowIndex) => _manualRowBreaks.Add(rowIndex);

    /// <summary>
    /// シートに画像(要件9)を1枚配置する。<paramref name="row"/>/<paramref name="column"/>のセル左上を
    /// 基準に、そこから<paramref name="offsetXPt"/>/<paramref name="offsetYPt"/>だけ離れた位置へ、
    /// <paramref name="widthPt"/>x<paramref name="heightPt"/>の固定サイズで配置する(oneCellAnchor)。
    /// </summary>
    public void SetImage(
        int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt, byte[] png) =>
        _image = (row, column, offsetXPt, offsetYPt, widthPt, heightPt, png);

    public void SetText(int row, int column, string? text, uint styleIndex = 0) =>
        _cells[Reference(row, column)] = (row, column, text, styleIndex, false);

    public void SetNumber(int row, int column, double value, uint styleIndex = 0) =>
        _cells[Reference(row, column)] = (row, column, value, styleIndex, true);

    /// <summary>書式だけを設定した空セル(罫線を引くために必要)。</summary>
    public void SetStyleOnly(int row, int column, uint styleIndex) =>
        _cells[Reference(row, column)] = (row, column, null, styleIndex, false);

    public static string Reference(int row, int column) => ColumnName(column) + row.ToString();

    public static string ColumnName(int column)
    {
        var name = string.Empty;
        var n = column;
        while (n > 0)
        {
            var rem = (n - 1) % 26;
            name = (char)('A' + rem) + name;
            n = (n - 1) / 26;
        }

        return name;
    }

    public void Save(string path, Stylesheet stylesheet)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = stylesheet;
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = BuildWorksheet();

        if (_image is { } image)
        {
            AppendImage(worksheetPart, image);
        }

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = SheetName,
        });

        AppendDefinedNames(workbookPart);

        workbookPart.Workbook.Save();
    }

    private Worksheet BuildWorksheet()
    {
        var worksheet = new Worksheet();

        worksheet.Append(new SheetProperties(new PageSetupProperties { FitToPage = false }));
        worksheet.Append(new SheetFormatProperties { DefaultRowHeight = 13.5D, DefaultColumnWidth = 8.43D });

        if (_columnWidths.Count > 0)
        {
            var columns = new Columns();
            foreach (var (column, width) in _columnWidths.OrderBy(kv => kv.Key))
            {
                columns.Append(new Column
                {
                    Min = (uint)column,
                    Max = (uint)column,
                    Width = width,
                    CustomWidth = true,
                });
            }

            worksheet.Append(columns);
        }

        worksheet.Append(BuildSheetData());

        if (_mergedRanges.Count > 0)
        {
            var mergeCells = new MergeCells { Count = (uint)_mergedRanges.Count };
            foreach (var range in _mergedRanges)
            {
                mergeCells.Append(new MergeCell { Reference = range });
            }

            worksheet.Append(mergeCells);
        }

        worksheet.Append(new PageMargins
        {
            Left = MarginLeftIn,
            Right = MarginRightIn,
            Top = MarginTopIn,
            Bottom = MarginBottomIn,
            Header = 0.3D,
            Footer = 0.3D,
        });

        worksheet.Append(new PageSetup
        {
            PaperSize = PaperSizeCode,
            Orientation = Orientation,
            Scale = ScalePercent,
            PageOrder = PageOrderValues.DownThenOver,
        });

        var headerFooter = BuildHeaderFooter();
        if (headerFooter is not null)
        {
            worksheet.Append(headerFooter);
        }

        if (_manualRowBreaks.Count > 0)
        {
            var rowBreaks = new RowBreaks
            {
                Count = (uint)_manualRowBreaks.Count,
                ManualBreakCount = (uint)_manualRowBreaks.Count,
            };

            foreach (var id in _manualRowBreaks)
            {
                rowBreaks.Append(new Break { Id = id, Max = 16383U, ManualPageBreak = true });
            }

            worksheet.Append(rowBreaks);
        }

        return worksheet;
    }

    private HeaderFooter? BuildHeaderFooter()
    {
        if (OddHeader is null && OddFooter is null && FirstHeader is null && FirstFooter is null)
        {
            return null;
        }

        var headerFooter = new HeaderFooter();
        if (DifferentFirstPage)
        {
            headerFooter.DifferentFirst = true;
        }

        if (OddHeader is not null)
        {
            headerFooter.Append(new OddHeader(OddHeader));
        }

        if (OddFooter is not null)
        {
            headerFooter.Append(new OddFooter(OddFooter));
        }

        if (FirstHeader is not null)
        {
            headerFooter.Append(new FirstHeader(FirstHeader));
        }

        if (FirstFooter is not null)
        {
            headerFooter.Append(new FirstFooter(FirstFooter));
        }

        return headerFooter;
    }

    private SheetData BuildSheetData()
    {
        var sheetData = new SheetData();

        var byRow = _cells.Values.GroupBy(c => c.Row).OrderBy(g => g.Key);
        foreach (var group in byRow)
        {
            var row = new Row { RowIndex = (uint)group.Key };
            if (_rowHeights.TryGetValue(group.Key, out var height))
            {
                row.Height = height;
                row.CustomHeight = true;
            }

            foreach (var entry in group.OrderBy(c => c.Column))
            {
                row.Append(BuildCell(entry));
            }

            sheetData.Append(row);
        }

        // 高さだけを指定した空行も出力する(行高が改ページ計算に影響するため)。
        foreach (var (rowIndex, height) in _rowHeights.OrderBy(kv => kv.Key))
        {
            if (byRow.Any(g => g.Key == rowIndex))
            {
                continue;
            }

            sheetData.Append(new Row { RowIndex = (uint)rowIndex, Height = height, CustomHeight = true });
        }

        // 行は昇順で並んでいる必要がある。
        var ordered = sheetData.Elements<Row>().OrderBy(r => r.RowIndex!.Value).ToList();
        sheetData.RemoveAllChildren();
        foreach (var row in ordered)
        {
            sheetData.Append(row);
        }

        return sheetData;
    }

    private static Cell BuildCell((int Row, int Column, object? Value, uint StyleIndex, bool IsNumber) entry)
    {
        var cell = new Cell
        {
            CellReference = Reference(entry.Row, entry.Column),
            StyleIndex = entry.StyleIndex,
        };

        if (entry.IsNumber && entry.Value is double number)
        {
            cell.CellValue = new CellValue(number);
            return cell;
        }

        if (entry.Value is string text && text.Length > 0)
        {
            // 共有文字列表を使わず inlineStr で埋め込む(サンプル生成を単純にするため)。
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text(text));
        }

        return cell;
    }

    private void AppendDefinedNames(WorkbookPart workbookPart)
    {
        if (PrintArea is null && PrintTitleRows is null)
        {
            return;
        }

        var definedNames = new DefinedNames();
        var quotedSheet = QuoteSheetName(SheetName);

        if (PrintArea is { } printArea)
        {
            // 複数の印刷範囲は、範囲ごとにシート名を付けてカンマで連結する
            // ("'納品書'!$A$1:$F$20,'納品書'!$A$22:$F$41" の形)。
            var qualified = string.Join(
                ",",
                printArea.Split(',').Select(range => $"{quotedSheet}!{range.Trim()}"));

            definedNames.Append(new DefinedName
            {
                Name = "_xlnm.Print_Area",
                LocalSheetId = 0U,
                Text = qualified,
            });
        }

        if (PrintTitleRows is { } titleRows)
        {
            definedNames.Append(new DefinedName
            {
                Name = "_xlnm.Print_Titles",
                LocalSheetId = 0U,
                Text = $"{quotedSheet}!{titleRows}",
            });
        }

        // definedNames は sheets より前に置く必要がある。
        workbookPart.Workbook.InsertBefore(definedNames, workbookPart.Workbook.Sheets);
    }

    private static string QuoteSheetName(string name) =>
        name.Any(c => char.IsWhiteSpace(c) || c > 0x7F) ? "'" + name.Replace("'", "''") + "'" : name;

    /// <summary>
    /// 画像パート(<c>xdr:pic</c>、oneCellAnchor)をワークシートへ追加する。
    /// </summary>
    private static void AppendImage(
        WorksheetPart worksheetPart,
        (int Row, int Column, double OffsetXPt, double OffsetYPt, double WidthPt, double HeightPt, byte[] Png) image)
    {
        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(image.Png))
        {
            imagePart.FeedData(stream);
        }

        var offsetXEmu = (long)Math.Round(image.OffsetXPt * EmusPerPoint);
        var offsetYEmu = (long)Math.Round(image.OffsetYPt * EmusPerPoint);
        var widthEmu = (long)Math.Round(image.WidthPt * EmusPerPoint);
        var heightEmu = (long)Math.Round(image.HeightPt * EmusPerPoint);

        var anchor = new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((image.Column - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset(offsetXEmu.ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowId((image.Row - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset(offsetYEmu.ToString(CultureInfo.InvariantCulture))),
            new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
            new Xdr.Picture(
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
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
            new Xdr.ClientData());

        var drawing = new Xdr.WorksheetDrawing();
        drawing.Append(anchor);
        drawingsPart.WorksheetDrawing = drawing;

        // CT_Worksheetのスキーマ順(drawingはrowBreaks/colBreaksより後)に従い、最後に追加する。
        worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
    }
}
}
