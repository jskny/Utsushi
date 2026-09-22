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
    private readonly List<ImageSpec> _images = new();
    private readonly List<ShapeSpec> _shapes = new();
    private readonly List<ConnectorSpec> _connectors = new();
    private readonly List<GroupSpec> _groups = new();

    /// <summary>画像(要件9)1枚ぶんの配置情報。</summary>
    private readonly struct ImageSpec
    {
        public ImageSpec(
            int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
            byte[] png, double rotationDegrees)
        {
            Row = row;
            Column = column;
            OffsetXPt = offsetXPt;
            OffsetYPt = offsetYPt;
            WidthPt = widthPt;
            HeightPt = heightPt;
            Png = png;
            RotationDegrees = rotationDegrees;
        }

        public int Row { get; }

        public int Column { get; }

        public double OffsetXPt { get; }

        public double OffsetYPt { get; }

        public double WidthPt { get; }

        public double HeightPt { get; }

        public byte[] Png { get; }

        public double RotationDegrees { get; }
    }

    /// <summary>図形(要件10)1つぶんの配置情報。</summary>
    private readonly struct ShapeSpec
    {
        public ShapeSpec(
            int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
            A.ShapeTypeValues preset, string? fillHex, string? outlineHex, double rotationDegrees, string? text,
            GradientSpec? gradient = null)
        {
            Row = row;
            Column = column;
            OffsetXPt = offsetXPt;
            OffsetYPt = offsetYPt;
            WidthPt = widthPt;
            HeightPt = heightPt;
            Preset = preset;
            FillHex = fillHex;
            OutlineHex = outlineHex;
            RotationDegrees = rotationDegrees;
            Text = text;
            Gradient = gradient;
        }

        public int Row { get; }

        public int Column { get; }

        public double OffsetXPt { get; }

        public double OffsetYPt { get; }

        public double WidthPt { get; }

        public double HeightPt { get; }

        public A.ShapeTypeValues Preset { get; }

        public string? FillHex { get; }

        public string? OutlineHex { get; }

        public double RotationDegrees { get; }

        public string? Text { get; }

        public GradientSpec? Gradient { get; }
    }

    /// <summary>グラデーション塗り(要件10.6)1つぶんの指定。</summary>
    internal readonly struct GradientSpec
    {
        public GradientSpec(IReadOnlyList<(double Position, string ColorHex)> stops, double angleDegrees, bool radial)
        {
            Stops = stops;
            AngleDegrees = angleDegrees;
            Radial = radial;
        }

        public IReadOnlyList<(double Position, string ColorHex)> Stops { get; }

        public double AngleDegrees { get; }

        public bool Radial { get; }
    }

    /// <summary>接続線(要件10.9)1つぶんの配置情報。</summary>
    private readonly struct ConnectorSpec
    {
        public ConnectorSpec(
            int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
            A.ShapeTypeValues preset, string? outlineHex, bool flipHorizontal, bool flipVertical,
            int? startShapeHandle, uint? startSiteIndex, int? endShapeHandle, uint? endSiteIndex)
        {
            Row = row;
            Column = column;
            OffsetXPt = offsetXPt;
            OffsetYPt = offsetYPt;
            WidthPt = widthPt;
            HeightPt = heightPt;
            Preset = preset;
            OutlineHex = outlineHex;
            FlipHorizontal = flipHorizontal;
            FlipVertical = flipVertical;
            StartShapeHandle = startShapeHandle;
            StartSiteIndex = startSiteIndex;
            EndShapeHandle = endShapeHandle;
            EndSiteIndex = endSiteIndex;
        }

        public int Row { get; }

        public int Column { get; }

        public double OffsetXPt { get; }

        public double OffsetYPt { get; }

        public double WidthPt { get; }

        public double HeightPt { get; }

        public A.ShapeTypeValues Preset { get; }

        public string? OutlineHex { get; }

        public bool FlipHorizontal { get; }

        public bool FlipVertical { get; }

        /// <summary>接続先の始点(要件10.11)。<see cref="SetShape"/>が返した図形のハンドル。<c>null</c>なら接続先無し。</summary>
        public int? StartShapeHandle { get; }

        public uint? StartSiteIndex { get; }

        /// <summary>接続先の終点(要件10.11)。<see cref="StartShapeHandle"/>と同様。</summary>
        public int? EndShapeHandle { get; }

        public uint? EndSiteIndex { get; }
    }

    /// <summary>グループ(要件10.10)内の図形子要素1つぶんの配置情報(子座標空間上、ポイント単位)。</summary>
    internal readonly struct GroupChildShapeSpec
    {
        public GroupChildShapeSpec(
            double offsetXPt, double offsetYPt, double widthPt, double heightPt,
            A.ShapeTypeValues preset, string? fillHex, string? outlineHex)
        {
            OffsetXPt = offsetXPt;
            OffsetYPt = offsetYPt;
            WidthPt = widthPt;
            HeightPt = heightPt;
            Preset = preset;
            FillHex = fillHex;
            OutlineHex = outlineHex;
        }

        public double OffsetXPt { get; }

        public double OffsetYPt { get; }

        public double WidthPt { get; }

        public double HeightPt { get; }

        public A.ShapeTypeValues Preset { get; }

        public string? FillHex { get; }

        public string? OutlineHex { get; }
    }

    /// <summary>グループ(要件10.10)1つぶんの配置情報。子座標空間は(0,0)を原点とし、グループ自身の表示サイズと同じ大きさとする(倍率1.0)。</summary>
    private readonly struct GroupSpec
    {
        public GroupSpec(
            int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
            double rotationDegrees, IReadOnlyList<GroupChildShapeSpec> children)
        {
            Row = row;
            Column = column;
            OffsetXPt = offsetXPt;
            OffsetYPt = offsetYPt;
            WidthPt = widthPt;
            HeightPt = heightPt;
            RotationDegrees = rotationDegrees;
            Children = children;
        }

        public int Row { get; }

        public int Column { get; }

        public double OffsetXPt { get; }

        public double OffsetYPt { get; }

        public double WidthPt { get; }

        public double HeightPt { get; }

        public double RotationDegrees { get; }

        public IReadOnlyList<GroupChildShapeSpec> Children { get; }
    }

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
    /// シートに画像(要件9)を1枚追加する。<paramref name="row"/>/<paramref name="column"/>のセル左上を
    /// 基準に、そこから<paramref name="offsetXPt"/>/<paramref name="offsetYPt"/>だけ離れた位置へ、
    /// <paramref name="widthPt"/>x<paramref name="heightPt"/>の固定サイズで配置する(oneCellAnchor)。
    /// 複数回呼べば出現順に重なる。
    /// </summary>
    /// <param name="rotationDegrees">回転角(度、時計回り。要件9.7)。</param>
    public void SetImage(
        int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt, byte[] png,
        double rotationDegrees = 0) =>
        _images.Add(new ImageSpec(row, column, offsetXPt, offsetYPt, widthPt, heightPt, png, rotationDegrees));

    /// <summary>
    /// シートに図形(要件10)を1つ追加する(oneCellAnchor)。複数回呼べば出現順に重なる。
    /// </summary>
    /// <param name="fillHex">塗りつぶし色(6桁16進、例: "1F4E8C")。nullは塗りつぶし無し。</param>
    /// <param name="outlineHex">枠線色(6桁16進)。nullは枠線無し。</param>
    /// <param name="rotationDegrees">回転角(度、時計回り)。</param>
    /// <param name="text">図形内テキスト。nullはテキスト無し。</param>
    /// <returns>
    /// この図形のハンドル(<see cref="SetConnector"/>の接続先指定に使う。要件10.11)。
    /// 実際の<c>NonVisualDrawingProperties/@id</c>とは異なり、このビルダー内でだけ意味を持つ通し番号。
    /// </returns>
    public int SetShape(
        int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
        A.ShapeTypeValues preset, string? fillHex, string? outlineHex, double rotationDegrees = 0, string? text = null)
    {
        var handle = _shapes.Count;
        _shapes.Add(new ShapeSpec(
            row, column, offsetXPt, offsetYPt, widthPt, heightPt, preset, fillHex, outlineHex, rotationDegrees, text));
        return handle;
    }

    /// <summary>
    /// シートに、グラデーション塗り(要件10.6)の図形を1つ追加する。<paramref name="gradientStops"/>は
    /// 位置(0.0〜1.0)と色(6桁16進)の組を2点以上指定する。<paramref name="radial"/>が
    /// <c>true</c>なら放射状、<c>false</c>なら<paramref name="angleDegrees"/>を角度とする線形グラデーション。
    /// </summary>
    public void SetGradientShape(
        int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
        A.ShapeTypeValues preset, IReadOnlyList<(double Position, string ColorHex)> gradientStops,
        double angleDegrees = 0, bool radial = false, string? outlineHex = null, double rotationDegrees = 0, string? text = null) =>
        _shapes.Add(new ShapeSpec(
            row, column, offsetXPt, offsetYPt, widthPt, heightPt, preset, fillHex: null, outlineHex, rotationDegrees, text,
            new GradientSpec(gradientStops, angleDegrees, radial)));

    /// <summary>
    /// シートに接続線(要件10.9)を1つ追加する(oneCellAnchor)。<paramref name="preset"/>には
    /// <c>straightConnector1</c>/<c>bentConnector2</c>/<c>bentConnector3</c>/
    /// <c>curvedConnector2</c>/<c>curvedConnector3</c>のいずれかを指定する。
    /// </summary>
    /// <param name="startShapeHandle">
    /// 始点の接続先(要件10.11)。<see cref="SetShape"/>が返したハンドル。<c>null</c>なら接続先無し
    /// (アンカー矩形の対角点をそのまま始点にする、要件10.9の既定動作)。
    /// </param>
    /// <param name="startSiteIndex">始点の接続点番号(<c>a:stCxn/@idx</c>。0=上,1=左,2=下,3=右)。</param>
    /// <param name="endShapeHandle">終点の接続先。<paramref name="startShapeHandle"/>と同様。</param>
    /// <param name="endSiteIndex">終点の接続点番号。<paramref name="startSiteIndex"/>と同様。</param>
    public void SetConnector(
        int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
        A.ShapeTypeValues preset, string? outlineHex = null, bool flipHorizontal = false, bool flipVertical = false,
        int? startShapeHandle = null, uint? startSiteIndex = null, int? endShapeHandle = null, uint? endSiteIndex = null) =>
        _connectors.Add(new ConnectorSpec(
            row, column, offsetXPt, offsetYPt, widthPt, heightPt, preset, outlineHex, flipHorizontal, flipVertical,
            startShapeHandle, startSiteIndex, endShapeHandle, endSiteIndex));

    /// <summary>
    /// シートにグループ化された図形(要件10.10)を1つ追加する(oneCellAnchor)。
    /// 子要素の座標(<see cref="GroupChildShapeSpec"/>)は、グループ自身の表示サイズを
    /// 子座標空間の大きさ(倍率1.0)として指定する。
    /// </summary>
    public void SetGroup(
        int row, int column, double offsetXPt, double offsetYPt, double widthPt, double heightPt,
        double rotationDegrees, IReadOnlyList<GroupChildShapeSpec> children) =>
        _groups.Add(new GroupSpec(row, column, offsetXPt, offsetYPt, widthPt, heightPt, rotationDegrees, children));

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

        if (_images.Count > 0 || _shapes.Count > 0 || _connectors.Count > 0 || _groups.Count > 0)
        {
            AppendDrawingObjects(worksheetPart, _images, _shapes, _connectors, _groups);
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
    /// 画像(<c>xdr:pic</c>)・図形(<c>xdr:sp</c>)・接続線(<c>xdr:cxnSp</c>)・グループ(<c>xdr:grpSp</c>)を、
    /// 1つの<c>DrawingsPart</c>にまとめて追加する(出現順=重なり順。要件10.3)。
    /// 画像→図形→接続線→グループの順に固定で並べる(サンプル生成専用の簡略化。
    /// 実際のExcelファイルはdrawing.xml内の任意の出現順を取りうるが、z-order自体の検証は
    /// Parsingレイヤーのユニットテストが個別のフィクスチャで行う)。
    /// </summary>
    private static void AppendDrawingObjects(
        WorksheetPart worksheetPart,
        IReadOnlyList<ImageSpec> images,
        IReadOnlyList<ShapeSpec> shapes,
        IReadOnlyList<ConnectorSpec> connectors,
        IReadOnlyList<GroupSpec> groups)
    {
        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        var drawing = new Xdr.WorksheetDrawing();

        var nextId = 2U;

        foreach (var image in images)
        {
            drawing.Append(BuildImageAnchor(drawingsPart, image, nextId++));
        }

        // 接続線の接続先解決(要件10.11)のため、SetShapeが返したハンドル(_shapesのindex)から
        // 実際に割り当てたIDへのマップを、ID割り当てと同じ順序で構築する。
        var shapeIds = new uint[shapes.Count];
        for (var i = 0; i < shapes.Count; i++)
        {
            shapeIds[i] = nextId;
            drawing.Append(BuildShapeAnchor(shapes[i], nextId));
            nextId++;
        }

        foreach (var connector in connectors)
        {
            var startShapeId = connector.StartShapeHandle is { } startHandle ? shapeIds[startHandle] : (uint?)null;
            var endShapeId = connector.EndShapeHandle is { } endHandle ? shapeIds[endHandle] : (uint?)null;
            drawing.Append(BuildConnectorAnchor(
                connector, nextId++, startShapeId, connector.StartSiteIndex, endShapeId, connector.EndSiteIndex));
        }

        foreach (var group in groups)
        {
            drawing.Append(BuildGroupAnchor(group, ref nextId));
        }

        drawingsPart.WorksheetDrawing = drawing;

        // CT_Worksheetのスキーマ順(drawingはrowBreaks/colBreaksより後)に従い、最後に追加する。
        worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
    }

    private static Xdr.OneCellAnchor BuildImageAnchor(DrawingsPart drawingsPart, ImageSpec image, uint id)
    {
        var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(image.Png))
        {
            imagePart.FeedData(stream);
        }

        var offsetXEmu = (long)Math.Round(image.OffsetXPt * EmusPerPoint);
        var offsetYEmu = (long)Math.Round(image.OffsetYPt * EmusPerPoint);
        var widthEmu = (long)Math.Round(image.WidthPt * EmusPerPoint);
        var heightEmu = (long)Math.Round(image.HeightPt * EmusPerPoint);
        var rotationEmu = (int)Math.Round(image.RotationDegrees * 60000.0);

        return new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((image.Column - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset(offsetXEmu.ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowId((image.Row - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset(offsetYEmu.ToString(CultureInfo.InvariantCulture))),
            new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
            new Xdr.Picture(
                new Xdr.NonVisualPictureProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = "Logo" },
                    new Xdr.NonVisualPictureDrawingProperties()),
                new Xdr.BlipFill(
                    new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart) },
                    new A.Stretch(new A.FillRectangle())),
                new Xdr.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = widthEmu, Cy = heightEmu })
                    {
                        Rotation = rotationEmu,
                    },
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
            new Xdr.ClientData());
    }

    /// <summary>図形(<c>xdr:sp</c>、oneCellAnchor)を組み立てる(要件10)。</summary>
    private static Xdr.OneCellAnchor BuildShapeAnchor(ShapeSpec shape, uint id)
    {
        var offsetXEmu = (long)Math.Round(shape.OffsetXPt * EmusPerPoint);
        var offsetYEmu = (long)Math.Round(shape.OffsetYPt * EmusPerPoint);
        var widthEmu = (long)Math.Round(shape.WidthPt * EmusPerPoint);
        var heightEmu = (long)Math.Round(shape.HeightPt * EmusPerPoint);
        var rotationEmu = (int)Math.Round(shape.RotationDegrees * 60000.0);

        OpenXmlElement fill = shape.Gradient is { } gradient
            ? BuildGradientFill(gradient)
            : shape.FillHex is { } fillHex
                ? new A.SolidFill(new A.RgbColorModelHex { Val = fillHex })
                : new A.NoFill();

        var shapeProperties = new Xdr.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = widthEmu, Cy = heightEmu })
            {
                Rotation = rotationEmu,
            },
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = shape.Preset },
            fill);

        if (shape.OutlineHex is { } outlineHex)
        {
            shapeProperties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = outlineHex })) { Width = 12700 });
        }

        var visualShape = new Xdr.Shape(
            new Xdr.NonVisualShapeProperties(
                new Xdr.NonVisualDrawingProperties { Id = id, Name = "Shape" + id.ToString(CultureInfo.InvariantCulture) },
                new Xdr.NonVisualShapeDrawingProperties()),
            shapeProperties);

        if (shape.Text is { } text)
        {
            visualShape.Append(new Xdr.TextBody(
                new A.BodyProperties { Anchor = A.TextAnchoringTypeValues.Center },
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Center },
                    new A.Run(
                        new A.RunProperties { FontSize = 1000 },
                        new A.Text(text)))));
        }

        return new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((shape.Column - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset(offsetXEmu.ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowId((shape.Row - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset(offsetYEmu.ToString(CultureInfo.InvariantCulture))),
            new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
            visualShape,
            new Xdr.ClientData());
    }

    /// <summary>グラデーション塗り(<c>a:gradFill</c>、要件10.6)を組み立てる。</summary>
    private static A.GradientFill BuildGradientFill(GradientSpec gradient)
    {
        var stopList = new A.GradientStopList();
        foreach (var (position, colorHex) in gradient.Stops)
        {
            var permille = (int)Math.Round(Math.Max(0.0, Math.Min(1.0, position)) * 100000.0);
            stopList.Append(new A.GradientStop(new A.RgbColorModelHex { Val = colorHex }) { Position = permille });
        }

        var gradientFill = new A.GradientFill(stopList);
        if (gradient.Radial)
        {
            gradientFill.Append(new A.PathGradientFill(new A.FillToRectangle()) { Path = A.PathShadeValues.Circle });
        }
        else
        {
            var angleEmu = (int)Math.Round(gradient.AngleDegrees * 60000.0);
            gradientFill.Append(new A.LinearGradientFill { Angle = angleEmu });
        }

        return gradientFill;
    }

    /// <summary>
    /// 接続線(<c>xdr:cxnSp</c>、oneCellAnchor)を組み立てる(要件10.9)。
    /// <paramref name="startShapeId"/>/<paramref name="endShapeId"/>が指定されていれば、
    /// <c>xdr:cNvCxnSpPr</c>配下に<c>a:stCxn</c>/<c>a:endCxn</c>(要件10.11)を追加する。
    /// </summary>
    private static Xdr.OneCellAnchor BuildConnectorAnchor(
        ConnectorSpec connector, uint id, uint? startShapeId, uint? startSiteIndex, uint? endShapeId, uint? endSiteIndex)
    {
        var offsetXEmu = (long)Math.Round(connector.OffsetXPt * EmusPerPoint);
        var offsetYEmu = (long)Math.Round(connector.OffsetYPt * EmusPerPoint);
        var widthEmu = (long)Math.Round(connector.WidthPt * EmusPerPoint);
        var heightEmu = (long)Math.Round(connector.HeightPt * EmusPerPoint);

        var transform = new A.Transform2D(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = widthEmu, Cy = heightEmu });
        if (connector.FlipHorizontal)
        {
            transform.HorizontalFlip = true;
        }

        if (connector.FlipVertical)
        {
            transform.VerticalFlip = true;
        }

        var shapeProperties = new Xdr.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = connector.Preset });

        if (connector.OutlineHex is { } outlineHex)
        {
            shapeProperties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = outlineHex })) { Width = 12700 });
        }

        var connectorShapeDrawingProperties = new Xdr.NonVisualConnectorShapeDrawingProperties();
        if (startShapeId is { } startId)
        {
            connectorShapeDrawingProperties.Append(new A.StartConnection { Id = startId, Index = startSiteIndex ?? 0 });
        }

        if (endShapeId is { } endId)
        {
            connectorShapeDrawingProperties.Append(new A.EndConnection { Id = endId, Index = endSiteIndex ?? 0 });
        }

        var connectionShape = new Xdr.ConnectionShape(
            new Xdr.NonVisualConnectionShapeProperties(
                new Xdr.NonVisualDrawingProperties { Id = id, Name = "Connector" + id.ToString(CultureInfo.InvariantCulture) },
                connectorShapeDrawingProperties),
            shapeProperties);

        return new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((connector.Column - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset(offsetXEmu.ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowId((connector.Row - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset(offsetYEmu.ToString(CultureInfo.InvariantCulture))),
            new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
            connectionShape,
            new Xdr.ClientData());
    }

    /// <summary>
    /// グループ(<c>xdr:grpSp</c>、oneCellAnchor)を組み立てる(要件10.10)。子座標空間は
    /// (0,0)を原点とし、グループ自身の表示サイズと同じ大きさ(倍率1.0)とする。
    /// </summary>
    private static Xdr.OneCellAnchor BuildGroupAnchor(GroupSpec group, ref uint nextId)
    {
        var offsetXEmu = (long)Math.Round(group.OffsetXPt * EmusPerPoint);
        var offsetYEmu = (long)Math.Round(group.OffsetYPt * EmusPerPoint);
        var widthEmu = (long)Math.Round(group.WidthPt * EmusPerPoint);
        var heightEmu = (long)Math.Round(group.HeightPt * EmusPerPoint);
        var rotationEmu = (int)Math.Round(group.RotationDegrees * 60000.0);

        var groupId = nextId++;
        var groupShapeProperties = new Xdr.GroupShapeProperties(
            new A.TransformGroup(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = widthEmu, Cy = heightEmu },
                new A.ChildOffset { X = 0L, Y = 0L },
                new A.ChildExtents { Cx = widthEmu, Cy = heightEmu })
            {
                Rotation = rotationEmu,
            });

        var groupShape = new Xdr.GroupShape(
            new Xdr.NonVisualGroupShapeProperties(
                new Xdr.NonVisualDrawingProperties { Id = groupId, Name = "Group" + groupId.ToString(CultureInfo.InvariantCulture) },
                new Xdr.NonVisualGroupShapeDrawingProperties()),
            groupShapeProperties);

        foreach (var child in group.Children)
        {
            groupShape.Append(BuildGroupChildShape(child, nextId++));
        }

        return new Xdr.OneCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId((group.Column - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset(offsetXEmu.ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowId((group.Row - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset(offsetYEmu.ToString(CultureInfo.InvariantCulture))),
            new Xdr.Extent { Cx = widthEmu, Cy = heightEmu },
            groupShape,
            new Xdr.ClientData());
    }

    /// <summary>グループ内の図形子要素(<c>xdr:sp</c>、子座標空間上の<c>a:off</c>/<c>a:ext</c>)を組み立てる。</summary>
    private static Xdr.Shape BuildGroupChildShape(GroupChildShapeSpec child, uint id)
    {
        var offsetXEmu = (long)Math.Round(child.OffsetXPt * EmusPerPoint);
        var offsetYEmu = (long)Math.Round(child.OffsetYPt * EmusPerPoint);
        var widthEmu = (long)Math.Round(child.WidthPt * EmusPerPoint);
        var heightEmu = (long)Math.Round(child.HeightPt * EmusPerPoint);

        OpenXmlElement fill = child.FillHex is { } fillHex
            ? new A.SolidFill(new A.RgbColorModelHex { Val = fillHex })
            : new A.NoFill();

        var shapeProperties = new Xdr.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = offsetXEmu, Y = offsetYEmu },
                new A.Extents { Cx = widthEmu, Cy = heightEmu }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = child.Preset },
            fill);

        if (child.OutlineHex is { } outlineHex)
        {
            shapeProperties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = outlineHex })) { Width = 12700 });
        }

        return new Xdr.Shape(
            new Xdr.NonVisualShapeProperties(
                new Xdr.NonVisualDrawingProperties { Id = id, Name = "GroupChild" + id.ToString(CultureInfo.InvariantCulture) },
                new Xdr.NonVisualShapeDrawingProperties()),
            shapeProperties);
    }
}
}
