using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Dr = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// <c>DocumentFormat.OpenXml</c>(Open XML SDK, MIT)を用いた <see cref="IWorkbookReader"/> の実装。
    /// </summary>
    /// <remarks>
    /// Office Interop / COM には一切依存しない(要件7.2)。読み取り専用で開くため、入力ファイルは変更されない。
    /// </remarks>
    public sealed class OpenXmlWorkbookReader : IWorkbookReader
    {
        /// <inheritdoc />
        public WorkbookModel Read(Stream xlsxStream, WorkbookReadOptions? options = null)
        {
            if (xlsxStream is null)
            {
                throw new ArgumentNullException(nameof(xlsxStream));
            }

            var opts = options ?? WorkbookReadOptions.Default;

            // SpreadsheetDocument はシーク可能なストリームを要求するため、必要ならメモリ上へ複製する。
            // 複製時のバイト数には上限を設ける(要件6, 7。シーク不可ストリーム経由での無制限な
            // メモリ確保を防ぐ安全弁。security-reviewer指摘)。
            var seekable = EnsureSeekable(xlsxStream, out var ownsStream, opts.ReportCode);
            try
            {
                return ReadCore(seekable, opts);
            }
            finally
            {
                if (ownsStream)
                {
                    seekable.Dispose();
                }
            }
        }

        /// <summary>ファイルパスからワークブックを読み取る補助メソッド。</summary>
        public WorkbookModel ReadFile(string path, WorkbookReadOptions? options = null)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("ファイルパスが空です。", nameof(path));
            }

            if (!File.Exists(path))
            {
                throw new InvalidExcelFileException(
                    $"入力ファイルが見つかりません: {path}",
                    InvalidExcelFileReason.Unknown,
                    options?.ReportCode);
            }

            using var stream = File.OpenRead(path);
            return Read(stream, options);
        }

        private static WorkbookModel ReadCore(Stream stream, WorkbookReadOptions options)
        {
            GuardPackageSize(stream, options.ReportCode);

            SpreadsheetDocument document;
            try
            {
                document = SpreadsheetDocument.Open(
                    stream,
                    isEditable: false,
                    new OpenSettings { MaxCharactersInPart = MaxXmlPartBytes });
            }
            catch (Exception ex)
            {
                throw MapOpenFailure(ex, stream, options.ReportCode);
            }

            using (document)
            {
                try
                {
                    return ReadDocument(document, options);
                }
                catch (Exception ex) when (ex is FormatException or OverflowException)
                {
                    // OpenXml SDK の型付き属性(.Value)は、不正な文字列(例: flipH="maybe"、r="abc")を読むと
                    // FormatException/OverflowException を投げる。個々の読み取り箇所で漏れなく扱うのは難しいため、
                    // ここでまとめて「壊れたファイル」として UtsushiException 階層に読み替える(要件6.4。security-reviewer指摘)。
                    throw new InvalidExcelFileException(
                        "入力ファイルに不正な値の属性が含まれているため読み取れません。",
                        InvalidExcelFileReason.Corrupted,
                        options.ReportCode,
                        ex);
                }
                catch (Exception ex) when (ex is InvalidDataException or XmlException or OpenXmlPackageException)
                {
                    // パートのルート要素が想定と違う(例: workbook.xml のルートが <foo/>)と、OpenXml SDK はパートを
                    // 初めて読む時点で InvalidDataException を投げる。壊れたXML・パッケージも同じく Corrupted に読み替える
                    // (要件6.4。security-reviewer指摘)。
                    throw new InvalidExcelFileException(
                        "入力ファイルの Open XML 構造が壊れているため読み取れません。",
                        InvalidExcelFileReason.Corrupted,
                        options.ReportCode,
                        ex);
                }
            }
        }

        private static WorkbookModel ReadDocument(SpreadsheetDocument document, WorkbookReadOptions options)
        {
            {
                // DOM に触れる前に、全XMLパートの入れ子の深さと大きさを流し読みで検査する(要件6.7)。
                GuardXmlParts(document, options.ReportCode);

                var workbookPart = document.WorkbookPart
                    ?? throw new InvalidExcelFileException(
                        "ワークブックパートが存在しません。ファイルが破損している可能性があります。",
                        InvalidExcelFileReason.Corrupted,
                        options.ReportCode);

                DetectUnsupportedWorkbookElements(workbookPart, options);

                var colors = ColorResolver.Create(workbookPart);
                var styles = StyleTable.Create(workbookPart, colors);
                var sharedStrings = ReadSharedStrings(workbookPart, options.ReportCode);
                var definedNames = ReadDefinedNames(workbookPart);
                var drawingColors = DrawingColorResolver.Create(workbookPart);

                // ActiveSheetOnly では、名前ではなく <sheet> 要素そのもので対象を決める。名前で絞ると、
                // 同じ名前の <sheet> が複数ある不正なファイルで別の要素(非表示・参照切れ)を読んでしまうため。
                X.Sheet? activeSheet = null;
                if (options.SheetNameFilter is null && options.ActiveSheetOnly)
                {
                    activeSheet = ResolveActiveSheet(workbookPart)
                        ?? throw new InvalidExcelFileException(
                            "表示されているワークシートが1つもありません(グラフシートは変換できません)。",
                            InvalidExcelFileReason.NoWorksheet,
                            options.ReportCode);
                }

                var sheets = new List<SheetModel>();
                var allSheetNames = new List<string>();
                foreach (var sheet in workbookPart.Workbook.Sheets?.Elements<X.Sheet>() ?? Enumerable.Empty<X.Sheet>())
                {
                    var name = sheet.Name?.Value;
                    if (string.IsNullOrEmpty(name) || sheet.Id?.Value is null)
                    {
                        continue;
                    }

                    allSheetNames.Add(name!);

                    if (options.SheetNameFilter is { } filter && !string.Equals(name, filter, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (activeSheet is not null && !ReferenceEquals(sheet, activeSheet))
                    {
                        continue;
                    }

                    if (GetSheetPart(workbookPart, sheet, options.ReportCode) is not WorksheetPart worksheetPart)
                    {
                        continue;
                    }

                    sheets.Add(ReadSheet(name!, worksheetPart, styles, sharedStrings, definedNames, drawingColors, options));

                    // 対象を1枚に絞っている場合は、見つけた時点で打ち切る。Excelはシート名の重複を許さないが、
                    // OpenXml SDK は検証しないため、同じ名前の <sheet> を大量に並べたファイルで同じシートを
                    // 何度も読まされるのを防ぐ(security-reviewer指摘)。
                    if (options.SheetNameFilter is not null || activeSheet is not null)
                    {
                        break;
                    }
                }

                if (sheets.Count == 0)
                {
                    if (options.SheetNameFilter is { } filter)
                    {
                        // フィルタで指定したシート名が一致しない場合。ファイル自体が壊れているのではなく
                        // 帳票定義が期待するシート構成と実際のブックが一致しないだけのため、原因を
                        // 特定できるよう実在するシート名を含める(呼び出し元がReportDefinition段階の
                        // エラーへ読み替える際に利用する)。
                        var available = string.Join(", ", allSheetNames.Select(n => "'" + n + "'"));
                        throw new InvalidExcelFileException(
                            $"シート '{filter}' が見つかりません。"
                            + $"ブック内のシート: {(available.Length == 0 ? "(なし)" : available)}",
                            InvalidExcelFileReason.NoWorksheet,
                            options.ReportCode);
                    }

                    throw new InvalidExcelFileException(
                        "ワークシートが1つも含まれていません。",
                        InvalidExcelFileReason.NoWorksheet,
                        options.ReportCode);
                }

                return new WorkbookModel(sheets, styles.DefaultFont);
            }
        }

        private static SheetModel ReadSheet(
            string name,
            WorksheetPart worksheetPart,
            StyleTable styles,
            IReadOnlyList<string> sharedStrings,
            IReadOnlyDictionary<string, Dictionary<string, string>> definedNames,
            DrawingColorResolver drawingColors,
            WorkbookReadOptions options)
        {
            var worksheet = worksheetPart.Worksheet;

            DetectUnsupportedElements(name, worksheetPart, worksheet, options);

            var cells = new Dictionary<CellAddress, CellModel>();
            var rowHeights = new List<double>();
            var hiddenRows = new HashSet<int>();

            var sheetFormat = worksheet.GetFirstChild<X.SheetFormatProperties>();
            // NaN・無限大・0以下・上限超えの既定行高は無視し、Excel の既定値にする(要件6.9)。
            var defaultRowHeight = ValidRowHeightOrNull(sheetFormat?.DefaultRowHeight?.Value) is { } validDefaultRowHeight
                && validDefaultRowHeight > 0.0
                    ? validDefaultRowHeight
                    : DefaultRowHeightPt;
            var defaultColumnWidth = ResolveDefaultColumnWidth(sheetFormat);

            var sheetData = worksheet.GetFirstChild<X.SheetData>();
            if (sheetData is not null)
            {
                var cellCount = 0;
                var previousRowIndex = 0L;
                foreach (var row in sheetData.Elements<X.Row>())
                {
                    // row/@r は省略できる(ECMA-376 Part 1, 18.3.1.73)。省略時は直前の行の次の行とする。
                    var rowIndexValue = row.RowIndex?.Value is { } explicitRow ? (long)explicitRow : previousRowIndex + 1;
                    if (rowIndexValue < 1)
                    {
                        continue;
                    }

                    // rowIndex自体がセル番地の上限近く(最大1,048,576)を指す不正な入力の場合、
                    // 実際のセル数に関わらずEnsureSizeが行高リストを巨大化させてしまうため、
                    // 個々のセルを読む前にこの時点で拒否する(要件6, 7。security-reviewer指摘)。
                    // 省略された r を補った番号にも同じ上限を当てる。
                    EnsureRowIndexWithinLimit((int)Math.Min(rowIndexValue, int.MaxValue), name, options.ReportCode);
                    var rowIndex = (int)rowIndexValue;
                    previousRowIndex = rowIndex;

                    EnsureSize(rowHeights, rowIndex, defaultRowHeight);

                    // Excel は自動調整された行(customHeight なし)にも ht を書くため、ht があれば customHeight によらず採用する。
                    // NaN・無限大・負・Excel の上限超えの値は無視し、既定の行高のままにする(要件6.9)。
                    if (ValidRowHeightOrNull(row.Height?.Value) is { } height)
                    {
                        rowHeights[rowIndex - 1] = height;
                    }

                    if (row.Hidden?.Value == true)
                    {
                        hiddenRows.Add(rowIndex);
                    }

                    var previousColumn = 0;
                    foreach (var cell in row.Elements<X.Cell>())
                    {
                        if (!TryResolveAddress(cell, rowIndex, previousColumn, out var address))
                        {
                            continue;
                        }

                        previousColumn = address.Column;
                        cellCount++;
                        EnsureCellCountWithinLimit(cellCount, name, options.ReportCode);

                        cells[address] = ReadCell(cell, styles, sharedStrings, name, options.ReportCode);
                    }
                }
            }

            var (columnWidths, hiddenColumns) = ReadColumns(worksheet, defaultColumnWidth, name, options.ReportCode);
            var mergedRanges = ReadMergedRanges(name, worksheet, options);
            var pageSetup = ReadPageSetup(name, worksheet, definedNames, options.ReportCode);
            var drawingObjects = ReadDrawingObjects(name, worksheetPart, drawingColors, options);

            return new SheetModel(
                name,
                cells,
                mergedRanges,
                columnWidths,
                rowHeights,
                defaultColumnWidth,
                defaultRowHeight,
                hiddenColumns,
                hiddenRows,
                pageSetup,
                drawingObjects,
                (int)Math.Min(sheetFormat?.BaseColumnWidth?.Value ?? 8U, 255U));
        }

        /// <summary>
        /// <c>defaultColWidth</c> が無いブックでは <c>baseColWidth</c>(既定8文字)から既定列幅を導く。
        /// </summary>
        private static double ResolveDefaultColumnWidth(X.SheetFormatProperties? sheetFormat)
        {
            // NaN・無限大・負・上限超えの defaultColWidth は無視し、baseColWidth から求める(要件6.9)。
            if (ValidColumnWidthOrNull(sheetFormat?.DefaultColumnWidth?.Value) is { } explicitWidth)
            {
                return explicitWidth;
            }

            // defaultColWidth が無い場合、Excel の既定列幅は baseColWidth と標準フォントの最大数字幅から決まる
            // (例: Calibri 11 で 64px、ＭＳ Ｐゴシック 11 で 72px)。最大数字幅は帳票定義側の値のため、ここでは
            // 「暗黙の既定幅」を NaN で表し、Layout レイヤーが求める(ExcelUnitConverter.DefaultColumnWidthToPixels)。
            return double.NaN;
        }

        private static CellModel ReadCell(
            X.Cell cell, StyleTable styles, IReadOnlyList<string> sharedStrings, string sheetName, string? reportCode)
        {
            var style = styles.GetCellStyle(cell.StyleIndex?.Value is { } s ? (int)s : null);
            var hasFormula = cell.CellFormula is not null;
            var rawValue = cell.CellValue?.InnerText;
            var dataType = cell.DataType?.Value;

            // インライン文字列
            if (dataType is not null && dataType == X.CellValues.InlineString)
            {
                // リッチテキストは共有文字列と同じく r/t だけを連結する。InnerText だと、ふりがな(rPh)の文字列まで
                // 本文に混ざる。
                var text = cell.InlineString is { } inline
                    ? inline.Text?.Text ?? string.Concat(inline.Elements<X.Run>().Select(r => r.Text?.Text ?? string.Empty))
                    : string.Empty;
                EnsureTextLengthWithinLimit(text, MaxCellTextLength, $"シート '{sheetName}' のセルの文字列", reportCode);
                return new CellModel(text, CellValueKind.Text, style, text, hasFormula);
            }

            // 共有文字列
            if (dataType is not null && dataType == X.CellValues.SharedString)
            {
                if (int.TryParse(rawValue, NumberStyles.None, CultureInfo.InvariantCulture, out var index)
                    && index >= 0 && index < sharedStrings.Count)
                {
                    var text = sharedStrings[index];
                    return new CellModel(text, CellValueKind.Text, style, text, hasFormula);
                }

                return new CellModel(null, CellValueKind.Blank, style, null, hasFormula);
            }

            if (string.IsNullOrEmpty(rawValue))
            {
                return new CellModel(null, CellValueKind.Blank, style, null, hasFormula);
            }

            if (dataType is not null && dataType == X.CellValues.Boolean)
            {
                var isTrue = rawValue == "1";
                return new CellModel(rawValue, CellValueKind.Boolean, style, isTrue ? "TRUE" : "FALSE", hasFormula);
            }

            if (dataType is not null && dataType == X.CellValues.Error)
            {
                return new CellModel(rawValue, CellValueKind.Error, style, rawValue, hasFormula);
            }

            if (dataType is not null && (dataType == X.CellValues.String || dataType == X.CellValues.Date))
            {
                // t="str" は数式の文字列結果、t="d" は ISO 8601 日付。いずれもそのまま表示する。
                return new CellModel(rawValue, CellValueKind.Text, style, rawValue, hasFormula);
            }

            // 型指定なし = 数値
            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            {
                var formatted = NumberFormatter.FormatNumber(number, style.NumberFormat);
                return new CellModel(rawValue, CellValueKind.Number, style, formatted, hasFormula);
            }

            return new CellModel(rawValue, CellValueKind.Text, style, rawValue, hasFormula);
        }

        private static (List<double> Widths, HashSet<int> Hidden) ReadColumns(
            X.Worksheet worksheet, double defaultWidth, string sheetName, string? reportCode)
        {
            var expansions = 0;

            var widths = new List<double>();
            var hidden = new HashSet<int>();

            var columns = worksheet.GetFirstChild<X.Columns>();
            if (columns is null)
            {
                return (widths, hidden);
            }

            foreach (var column in columns.Elements<X.Column>())
            {
                // uint のまま最大列で丸めてから int にする(int の範囲を超える値がキャストで負にならないように)。
                var min = (int)Math.Min(column.Min?.Value ?? 0U, (uint)CellAddress.MaxColumn + 1);
                var max = (int)Math.Min(column.Max?.Value ?? 0U, (uint)CellAddress.MaxColumn);
                if (min < 1)
                {
                    continue;
                }

                // Excel は未使用の右端まで Column 要素を伸ばすことがある。最大列を超える部分は丸めた。

                // min が最大列を超える定義は範囲が空になる。そのまま数えると展開回数の合計が負になり、
                // 上限(MaxColumnExpansionsPerSheet)をすり抜けられるため読み飛ばす(spec-compliance-reviewer指摘)。
                if (max < min)
                {
                    continue;
                }

                expansions += max - min + 1;
                EnsureColumnExpansionsWithinLimit(expansions, MaxColumnExpansionsPerSheet, sheetName, reportCode);

                var isHidden = column.Hidden?.Value == true;
                // NaN・無限大・負・上限超えの幅は、幅の指定が無いものとして扱う(要件6.9)。
                var explicitWidth = ValidColumnWidthOrNull(column.Width?.Value);
                var width = explicitWidth ?? defaultWidth;
                var hasCustomWidth = explicitWidth is not null;

                for (var c = min; c <= max; c++)
                {
                    EnsureSize(widths, c, defaultWidth);
                    if (hasCustomWidth)
                    {
                        widths[c - 1] = isHidden ? 0.0 : width;
                    }
                    else if (isHidden)
                    {
                        widths[c - 1] = 0.0;
                    }

                    if (isHidden)
                    {
                        hidden.Add(c);
                    }
                }
            }

            return (widths, hidden);
        }

        /// <summary>
        /// <c>mergeCell</c>要素を<see cref="MergedRange"/>のリストとして読み取る。個数には
        /// <see cref="MaxMergedRangesPerSheet"/>の上限を設ける(security-reviewer指摘。
        /// 画像・図形と同じ理由によるDoS対策)。
        /// </summary>
        private static List<MergedRange> ReadMergedRanges(string sheetName, X.Worksheet worksheet, WorkbookReadOptions options)
        {
            var result = new List<MergedRange>();
            var mergeCells = worksheet.GetFirstChild<X.MergeCells>();
            if (mergeCells is null)
            {
                return result;
            }

            foreach (var merge in mergeCells.Elements<X.MergeCell>())
            {
                if (merge.Reference?.Value is not { } reference || !CellRange.TryParse(reference, out var range))
                {
                    continue;
                }

                if (result.Count >= MaxMergedRangesPerSheet)
                {
                    if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                    {
                        throw new UnsupportedWorkbookElementException(
                            $"シート '{sheetName}' の結合セル範囲の数が上限({MaxMergedRangesPerSheet}個)を超えています。"
                            + "帳票定義の unsupportedElements が 'error' のため中止します。",
                            "TooManyMergedRanges",
                            options.ReportCode,
                            sheetName);
                    }

                    // ignore時は上限以降の結合範囲を無視する(以降のセルは通常セルとして扱われる)。
                    break;
                }

                result.Add(new MergedRange(range));
            }

            return result;
        }

        /// <summary>対応するラスター画像のMIMEタイプ(要件9.4)。SkiaSharpによる実デコードではなく、
        /// この静的な許可リストで判定する(ParsingレイヤーはRenderingレイヤーのSkiaSharpに依存しないため)。</summary>
        private static readonly HashSet<string> SupportedImageContentTypes = new(StringComparer.OrdinalIgnoreCase)
        {
            "image/png",
            "image/jpeg",
            "image/gif",
            "image/bmp",
        };

        /// <summary>
        /// 1シートに含める画像アンカーの数の上限。信頼できない入力(帳票定義への登録前の
        /// Excelファイル)による計算量の増大を防ぐための安全弁であり、自社ロゴ用途で
        /// この上限に達することは想定していない(security-reviewer指摘)。
        /// </summary>
        internal const int MaxImagesPerSheet = 50;

        /// <summary>画像1枚あたりの読み取りバイト数の上限(10MB)。ピクセル爆弾等への安全弁。</summary>
        private const long MaxImageDataBytes = 10 * 1024 * 1024;

        /// <summary>
        /// 1シートに含める結合セル範囲(<c>mergeCell</c>)の数の上限。画像・図形と同様、
        /// 信頼できない入力による計算量の増大を防ぐための安全弁である
        /// (Layoutレイヤーの結合セル矩形統合・罫線合成は結合範囲の個数に比例する処理のため。
        /// security-reviewer指摘)。個々の結合範囲の大きさ(行数・列数)自体は
        /// Layoutレイヤー側の<c>MaxSpanCells</c>で別途上限を設けている。
        /// </summary>
        internal const int MaxMergedRangesPerSheet = 1000;

        /// <summary>
        /// セルの文字列(共有文字列・インライン文字列)の長さの上限(要件6.8)。Excel 自体の上限(32,767文字)に合わせる。
        /// 極端に長い文字列を多数のセルから参照させ、レイアウト計算(文字幅の計測・折り返し)を長時間かけさせるのを防ぐ
        /// (security-reviewer指摘)。
        /// </summary>
        internal const int MaxCellTextLength = 32767;

        /// <summary>
        /// ヘッダー/フッター1つの文字列(書式コードを含む)の長さの上限(要件6.8)。Excel の画面では255文字までだが、
        /// 他のツールで作ったファイルを考慮して余裕をもたせる。
        /// </summary>
        internal const int MaxHeaderFooterTextLength = 1024;

        /// <summary>拡大縮小率(<c>pageSetup/@scale</c>)の範囲(Excel と同じ 10〜400%。要件6.9)。</summary>
        internal const int MinPrintScalePercent = 10;

        /// <summary>拡大縮小率の上限(%)。</summary>
        internal const int MaxPrintScalePercent = 400;

        /// <summary>既定の行高(pt)。<c>sheetFormatPr/@defaultRowHeight</c> が無いか不正な場合に使う。</summary>
        internal const double DefaultRowHeightPt = 15.0;

        /// <summary>行の高さの上限(pt。Excel 自体の上限。要件6.9)。これを超える値・NaN・無限大・負の値は無視する。</summary>
        internal const double MaxRowHeightPt = 409.0;

        /// <summary>列幅の上限(文字数。Excel 自体の上限。要件6.9)。これを超える値・NaN・無限大・負の値は無視する。</summary>
        internal const double MaxColumnWidthChars = 255.0;

        /// <summary>
        /// 余白1辺の上限(インチ。要件6.9)。どの用紙よりも大きい値とし、これを超える値・NaN・無限大・負の値は
        /// その辺の既定の余白に戻す。用紙に収まらない余白の検出は Layout レイヤーが行う。
        /// </summary>
        internal const double MaxPageMarginInches = 100.0;

        /// <summary>フォントサイズの下限(pt。Excel と同じ。要件6.9)。</summary>
        internal const double MinFontSizePt = 1.0;

        /// <summary>フォントサイズの上限(pt。Excel と同じ。要件6.9)。範囲外・NaN・無限大は既定のサイズに戻す。</summary>
        internal const double MaxFontSizePt = 409.0;

        /// <summary>図形の文字のサイズ(<c>a:rPr/@sz</c>。1/100pt)の下限。ECMA-376 の ST_TextFontSize(100〜400000)に合わせる。</summary>
        internal const int MinShapeFontSizeHundredths = 100;

        /// <summary>図形の文字のサイズ(<c>a:rPr/@sz</c>。1/100pt)の上限。</summary>
        internal const int MaxShapeFontSizeHundredths = 400_000;

        /// <summary>
        /// XMLパート1つに含める要素の数の上限(要件6.7)。パートの大きさの上限だけでは、小さな要素(例: <c>&lt;row/&gt;</c>)を
        /// 大量に並べて DOM に数GBのメモリを使わせられるため(100MiB 分の <c>&lt;row/&gt;</c> で 4.6GB。security-reviewer指摘)。
        /// <see cref="MaxCellsPerSheet"/> いっぱいのシート(セル1つあたり数要素)でも収まる値にする。
        /// </summary>
        internal const long MaxXmlElementsPerPart = 5_000_000;

        /// <summary>
        /// 関係パート(<c>*.rels</c>)1つに含める関係(<c>Relationship</c>)の数の上限(要件6.7)。関係パートは
        /// <c>SpreadsheetDocument.Open</c> の中で解析され、関係の数に対して2乗より速く処理時間が増える
        /// (10万件で139秒。security-reviewer指摘)ため、Open より前に ZIP を直接流し読みして検査する。
        /// 自社帳票の関係の数は、ハイパーリンクを多用しても数百件程度である。
        /// </summary>
        internal const int MaxRelationshipsPerPart = 10_000;

        /// <summary>
        /// パッケージ内の全関係パートの関係の数の合計の上限(要件6.7)。パートごとの上限(<see cref="MaxRelationshipsPerPart"/>)
        /// の直前まで詰めた関係パートを多数並べ、Open の処理時間を積み上げるのを防ぐ。
        /// </summary>
        internal const int MaxRelationshipsPerPackage = 50_000;

        /// <summary>
        /// ZIPエントリの数の上限(要件6.7)。ZipArchive はエントリごとにオブジェクトを作り、Open もパートごとに処理するため。
        /// 自社帳票のエントリは画像を含めても数十個程度である。
        /// </summary>
        internal const int MaxZipEntries = 10_000;

        /// <summary><c>[Content_Types].xml</c> のエントリ名。</summary>
        private const string ContentTypesEntryName = "[Content_Types].xml";

        /// <summary>
        /// <c>[Content_Types].xml</c> の展開後の大きさの上限(4MiB。要件6.7)。パートではないため <see cref="GuardXmlParts(SpreadsheetDocument, string?)"/>
        /// の対象にならず、Open の中で DOM として読まれる(100万件の Override で処理時間・メモリが極端に増えた。security-reviewer指摘)。
        /// </summary>
        internal const long MaxContentTypesBytes = 4L * 1024 * 1024;

        /// <summary><c>[Content_Types].xml</c> の要素(<c>Default</c>/<c>Override</c>)の数の合計の上限(要件6.7)。</summary>
        internal const int MaxContentTypesEntries = 10_000;

        /// <summary>
        /// パッケージ全体の XML パートの要素の数の合計の上限(要件6.7)。パートごとの上限(<see cref="MaxXmlElementsPerPart"/>)
        /// の直前まで詰めたパートを複数並べ、DOM のメモリを積み上げるのを防ぐ。
        /// </summary>
        internal const long MaxXmlElementsPerPackage = 8_000_000;

        /// <summary>
        /// XMLパートの要素の入れ子の深さの上限(要件6.7)。OpenXml SDK は DOM を再帰で組み立てるため、
        /// 数千段にネストした要素(例: <c>xdr:grpSp</c>)でスタックオーバーフローし、呼び出し元のプロセスごと
        /// 落ちる(スタック1MBで深さ5,000段・4.2KBのファイルで再現。security-reviewer指摘)。
        /// Excelが保存する帳票の入れ子は、グループを上限(<see cref="MaxShapeNestingDepth"/>)まで重ねても
        /// 数十段に収まるため、十分大きな値にする。
        /// </summary>
        internal const int MaxXmlElementDepth = 256;

        /// <summary>
        /// XMLパート1つの展開後の大きさの上限(64MiB。要件6.7)。<see cref="MaxCellsPerSheet"/>いっぱいの
        /// シートでも数十MBに収まる。<c>OpenSettings.MaxCharactersInPart</c>にも同じ値を設定する。
        /// </summary>
        internal const long MaxXmlPartBytes = 64L * 1024 * 1024;

        /// <summary>
        /// 1シートの列定義(<c>&lt;col&gt;</c>)の <c>min</c>〜<c>max</c> を展開する回数の合計の上限(要件6.8)。
        /// Excelは重ならない範囲で保存するため合計は最大列数(16,384)以下になる。範囲の重なる
        /// <c>&lt;col&gt;</c>を大量に並べた入力で処理時間が極端に増えるのを防ぐ(security-reviewer指摘)。
        /// </summary>
        internal const int MaxColumnExpansionsPerSheet = 4 * CellAddress.MaxColumn;

        /// <summary>手動改ページ(<c>&lt;brk&gt;</c>)の件数の上限(行・列それぞれ。Excel自体の上限。要件6.8)。</summary>
        internal const int MaxPageBreaksPerSheet = 1026;

        /// <summary>
        /// 1シートの印刷範囲(<c>_xlnm.Print_Area</c>のカンマ区切り)の個数の上限(要件6.8)。印刷範囲は
        /// 1個ごとに独立したページ群になるため、大量の範囲で数十万ページのPDFを出力させられるのを防ぐ
        /// (8.6KBのファイルで30万ページになった。security-reviewer指摘)。
        /// </summary>
        internal const int MaxPrintAreasPerSheet = 1000;

        /// <summary>
        /// 入力ファイル自体のバイト数、および展開後のZIPエントリ宣言サイズ合計の上限(1GiB)。
        /// 数百バイトのファイルが展開後に極端に大きくなる「ZIP爆弾」や、シーク不可ストリーム経由での
        /// 無制限なメモリ確保を防ぐための安全弁(要件6, 7。security-reviewer指摘)。自社帳票は
        /// 埋め込み画像を含めてもこの上限に達することは想定していない
        /// (画像1枚10MB×上限50枚=1シートあたり最大500MB相当が既存の上限だが、
        /// これより十分大きく設定し、通常の帳票運用を妨げないようにする)。
        /// </summary>
        internal const long MaxXlsxPackageBytes = 1024L * 1024 * 1024;

        /// <summary>
        /// 1ブックに含める共有文字列(<c>sharedStrings.xml</c>の<c>si</c>要素)の数の上限。
        /// 上限を超えて読み取ると、以降の共有文字列を参照するセルは(既存の「範囲外索引は空欄」
        /// 挙動と同じ経路で)空欄になる。画像・図形と異なり`unsupportedElements`の対象ではなく、
        /// 常にこの挙動になる(要件6, 7。security-reviewer指摘)。
        /// </summary>
        private const int MaxSharedStringCount = 200_000;

        /// <summary>
        /// 1シートに含めるセルの総数の上限。Excelの理論上限(最大1,048,576行×16,384列)まで
        /// 密にセルを敷き詰めた入力による計算量・メモリの増大を防ぐ安全弁(要件6, 7。
        /// security-reviewer指摘)。登録済み自社帳票は台帳形式で通常数百〜数千セル程度であり、
        /// この上限に達することは想定していない。
        /// </summary>
        private const int MaxCellsPerSheet = 500_000;

        /// <summary>
        /// <paramref name="rowIndex"/>が<see cref="MaxCellsPerSheet"/>を超えていないか確認する。
        /// 実際の上限は現実的なユニットテストでは大量の行を用意しないと到達できないため、
        /// 比較・例外構築のロジック自体を分離してテスト可能にしている。
        /// </summary>
        private static void EnsureRowIndexWithinLimit(int rowIndex, string sheetName, string? reportCode) =>
            EnsureRowIndexWithinLimit(rowIndex, MaxCellsPerSheet, sheetName, reportCode);

        internal static void EnsureRowIndexWithinLimit(int rowIndex, int maxRows, string sheetName, string? reportCode)
        {
            if (rowIndex > maxRows)
            {
                throw new InvalidExcelFileException(
                    $"シート '{sheetName}' の行番号({rowIndex})が上限({maxRows})を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }
        }

        /// <summary>
        /// アクティブシート(<c>workbookView/@activeTab</c>。省略時は0)の <c>sheet</c> 要素を返す。そのシートが非表示
        /// (<c>state="hidden"</c>/<c>"veryHidden"</c>)またはワークシートでない(グラフシート等)場合は、
        /// 表示されている最初のワークシートの要素を返す。該当するシートが無ければ null(要件12.2)。
        /// </summary>
        /// <remarks>
        /// 候補の判定条件は、本体の読み取りループ(<see cref="WorksheetPart"/>だけを読む)と揃えておくこと。
        /// ずれると、選んだシートが読み取りループで読み飛ばされて0枚になる。
        /// </remarks>
        private static X.Sheet? ResolveActiveSheet(WorkbookPart workbookPart)
        {
            var workbook = workbookPart.Workbook;
            var allSheets = workbook.Sheets?.Elements<X.Sheet>().ToList() ?? new List<X.Sheet>();

            bool IsCandidate(X.Sheet sheet) => IsVisibleSheet(sheet) && IsWorksheet(workbookPart, sheet);

            var activeTab = workbook.BookViews?.GetFirstChild<X.WorkbookView>()?.ActiveTab?.Value ?? 0U;
            if (activeTab < allSheets.Count && IsCandidate(allSheets[(int)activeTab]))
            {
                return allSheets[(int)activeTab];
            }

            return allSheets.FirstOrDefault(IsCandidate);
        }

        private static bool IsVisibleSheet(X.Sheet sheet) =>
            !string.IsNullOrEmpty(sheet.Name?.Value)
            && sheet.Id?.Value is not null
            && sheet.State?.Value != X.SheetStateValues.Hidden
            && sheet.State?.Value != X.SheetStateValues.VeryHidden;

        private static bool IsWorksheet(WorkbookPart workbookPart, X.Sheet sheet)
        {
            try
            {
                return workbookPart.GetPartById(sheet.Id!.Value!) is WorksheetPart;
            }
            catch (ArgumentOutOfRangeException)
            {
                // r:id に対応するパートが無い壊れた参照は候補にしない。
                return false;
            }
        }

        /// <summary>
        /// <c>sheet/@r:id</c> が指すパートを返す。参照先が無い壊れたファイルでは、OpenXml SDK の
        /// <see cref="ArgumentOutOfRangeException"/> を <see cref="InvalidExcelFileException"/> に読み替える
        /// (要件6.4: 失敗は <see cref="UtsushiException"/> 階層で返す。security-reviewer指摘)。
        /// </summary>
        private static OpenXmlPart GetSheetPart(WorkbookPart workbookPart, X.Sheet sheet, string? reportCode)
        {
            try
            {
                return workbookPart.GetPartById(sheet.Id!.Value!);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new InvalidExcelFileException(
                    $"シート '{sheet.Name?.Value}' の参照先パート('{sheet.Id?.Value}')がファイル内にありません。",
                    InvalidExcelFileReason.Corrupted,
                    reportCode,
                    ex);
            }
        }

        /// <summary>
        /// <paramref name="cellCount"/>が<see cref="MaxCellsPerSheet"/>を超えていないか確認する。
        /// <see cref="EnsureRowIndexWithinLimit(int, string, string?)"/>と同じ理由でテスト可能にしている。
        /// </summary>
        private static void EnsureCellCountWithinLimit(int cellCount, string sheetName, string? reportCode) =>
            EnsureCellCountWithinLimit(cellCount, MaxCellsPerSheet, sheetName, reportCode);

        /// <summary>文字列の長さが上限以下か確認する(要件6.8)。</summary>
        internal static void EnsureTextLengthWithinLimit(string text, int maxLength, string description, string? reportCode)
        {
            if (text.Length > maxLength)
            {
                throw new InvalidExcelFileException(
                    $"{description}の長さが上限({maxLength}文字)を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }
        }

        /// <summary><c>&lt;col&gt;</c>の展開回数の合計が上限以下か確認する(要件6.8。テスト用に上限を引数に取る)。</summary>
        internal static void EnsureColumnExpansionsWithinLimit(int expansions, int maxExpansions, string sheetName, string? reportCode)
        {
            if (expansions > maxExpansions)
            {
                throw new InvalidExcelFileException(
                    $"シート '{sheetName}' の列定義(<col>)が上限({maxExpansions}列分)を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }
        }

        internal static void EnsureCellCountWithinLimit(int cellCount, int maxCells, string sheetName, string? reportCode)
        {
            if (cellCount > maxCells)
            {
                throw new InvalidExcelFileException(
                    $"シート '{sheetName}' のセル数が上限({maxCells}個)を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }
        }

        /// <summary>
        /// 対応済みプリセットジオメトリ(要件10.1補足)。自社帳票での実用上の必要性を踏まえた
        /// キュレーション方式であり、ECMA-376の <c>ST_ShapeType</c> 全体には対応しない。
        /// </summary>
        private static readonly Dictionary<Dr.ShapeTypeValues, ShapePresetType> SupportedShapePresets = new()
        {
            [Dr.ShapeTypeValues.Rectangle] = ShapePresetType.Rect,
            [Dr.ShapeTypeValues.RoundRectangle] = ShapePresetType.RoundRect,
            [Dr.ShapeTypeValues.Ellipse] = ShapePresetType.Ellipse,
            [Dr.ShapeTypeValues.Triangle] = ShapePresetType.Triangle,
            [Dr.ShapeTypeValues.RightArrow] = ShapePresetType.RightArrow,
            [Dr.ShapeTypeValues.LeftArrow] = ShapePresetType.LeftArrow,
            [Dr.ShapeTypeValues.UpArrow] = ShapePresetType.UpArrow,
            [Dr.ShapeTypeValues.DownArrow] = ShapePresetType.DownArrow,
            [Dr.ShapeTypeValues.LeftRightArrow] = ShapePresetType.LeftRightArrow,
            [Dr.ShapeTypeValues.UpDownArrow] = ShapePresetType.UpDownArrow,
            [Dr.ShapeTypeValues.WedgeRectangleCallout] = ShapePresetType.WedgeRectCallout,
            [Dr.ShapeTypeValues.WedgeRoundRectangleCallout] = ShapePresetType.WedgeRoundRectCallout,
            [Dr.ShapeTypeValues.WedgeEllipseCallout] = ShapePresetType.WedgeEllipseCallout,
            [Dr.ShapeTypeValues.CloudCallout] = ShapePresetType.CloudCallout,
            [Dr.ShapeTypeValues.Callout1] = ShapePresetType.Callout1,
            [Dr.ShapeTypeValues.Callout2] = ShapePresetType.Callout2,
            [Dr.ShapeTypeValues.Callout3] = ShapePresetType.Callout3,
            [Dr.ShapeTypeValues.BorderCallout1] = ShapePresetType.BorderCallout1,
            [Dr.ShapeTypeValues.BorderCallout2] = ShapePresetType.BorderCallout2,
            [Dr.ShapeTypeValues.BorderCallout3] = ShapePresetType.BorderCallout3,
            [Dr.ShapeTypeValues.AccentCallout1] = ShapePresetType.AccentCallout1,
            [Dr.ShapeTypeValues.AccentCallout2] = ShapePresetType.AccentCallout2,
            [Dr.ShapeTypeValues.AccentCallout3] = ShapePresetType.AccentCallout3,
            [Dr.ShapeTypeValues.AccentBorderCallout1] = ShapePresetType.AccentBorderCallout1,
            [Dr.ShapeTypeValues.AccentBorderCallout2] = ShapePresetType.AccentBorderCallout2,
            [Dr.ShapeTypeValues.AccentBorderCallout3] = ShapePresetType.AccentBorderCallout3,
            [Dr.ShapeTypeValues.Star4] = ShapePresetType.Star4,
            [Dr.ShapeTypeValues.Star5] = ShapePresetType.Star5,
            [Dr.ShapeTypeValues.Star6] = ShapePresetType.Star6,
            [Dr.ShapeTypeValues.Star8] = ShapePresetType.Star8,
            [Dr.ShapeTypeValues.FlowChartProcess] = ShapePresetType.FlowChartProcess,
            [Dr.ShapeTypeValues.FlowChartDecision] = ShapePresetType.FlowChartDecision,
            [Dr.ShapeTypeValues.FlowChartTerminator] = ShapePresetType.FlowChartTerminator,
            [Dr.ShapeTypeValues.FlowChartInputOutput] = ShapePresetType.FlowChartInputOutput,
            [Dr.ShapeTypeValues.FlowChartDocument] = ShapePresetType.FlowChartDocument,
            [Dr.ShapeTypeValues.FlowChartPredefinedProcess] = ShapePresetType.FlowChartPredefinedProcess,
            [Dr.ShapeTypeValues.FlowChartConnector] = ShapePresetType.FlowChartConnector,
        };

        /// <summary>
        /// 対応済み接続線(<c>xdr:cxnSp</c>)プリセット一覧(要件10.9)。同じくキュレーション方式。
        /// </summary>
        private static readonly Dictionary<Dr.ShapeTypeValues, ConnectorPresetType> SupportedConnectorPresets = new()
        {
            [Dr.ShapeTypeValues.StraightConnector1] = ConnectorPresetType.Straight,

            // Excelの「図形 → 線 → 直線」は prst="line" の接続線として保存される。形は straightConnector1 と同じ
            // (アンカー矩形の対角を結ぶ直線。反転で向きが決まる)ため、同じ直線として扱う。
            [Dr.ShapeTypeValues.Line] = ConnectorPresetType.Straight,
            [Dr.ShapeTypeValues.BentConnector2] = ConnectorPresetType.Bent2Segment,
            [Dr.ShapeTypeValues.BentConnector3] = ConnectorPresetType.Bent3Segment,
            [Dr.ShapeTypeValues.CurvedConnector2] = ConnectorPresetType.Curved2Segment,
            [Dr.ShapeTypeValues.CurvedConnector3] = ConnectorPresetType.Curved3Segment,
        };

        private static readonly string[] LineCalloutGuideNames1 = { "adj1", "adj2", "adj3", "adj4" };

        private static readonly string[] LineCalloutGuideNames2 = { "adj1", "adj2", "adj3", "adj4", "adj5", "adj6" };

        private static readonly string[] LineCalloutGuideNames3 = { "adj1", "adj2", "adj3", "adj4", "adj5", "adj6", "adj7", "adj8" };

        /// <summary>
        /// プリセットごとの調整ガイド(<c>a:avLst/a:gd/@name</c>)の並び順。<see cref="ShapeModel.AdjustmentValues"/>
        /// はこの並びに対応する固定長のリストとして返し、ファイルにガイドが無い位置は
        /// <see cref="double.NaN"/> とする(Renderingレイヤーが該当位置のECMA-376既定値を補う)。
        /// </summary>
        private static readonly Dictionary<ShapePresetType, string[]> ShapeAdjustmentGuideNames = new()
        {
            [ShapePresetType.Rect] = Array.Empty<string>(),
            [ShapePresetType.Ellipse] = Array.Empty<string>(),
            [ShapePresetType.RoundRect] = new[] { "adj" },
            [ShapePresetType.Triangle] = new[] { "adj" },
            [ShapePresetType.RightArrow] = new[] { "adj1", "adj2" },
            [ShapePresetType.LeftArrow] = new[] { "adj1", "adj2" },
            [ShapePresetType.UpArrow] = new[] { "adj1", "adj2" },
            [ShapePresetType.DownArrow] = new[] { "adj1", "adj2" },
            [ShapePresetType.LeftRightArrow] = new[] { "adj1", "adj2" },
            [ShapePresetType.UpDownArrow] = new[] { "adj1", "adj2" },
            [ShapePresetType.WedgeRectCallout] = new[] { "adj1", "adj2" },
            [ShapePresetType.WedgeRoundRectCallout] = new[] { "adj1", "adj2" },
            [ShapePresetType.WedgeEllipseCallout] = new[] { "adj1", "adj2" },
            // cloudCallout本体の輪郭(バンプの個数・半径)は固定形状として近似描画するが、
            // 引き出し三角形の位置(adj1=X方向, adj2=Y方向)はwedgeRectCallout等と同じ意味の
            // 調整ガイドのため読み取る(要件10.13)。
            [ShapePresetType.CloudCallout] = new[] { "adj1", "adj2" },
            // 線吹き出し(callout/borderCallout/accentCallout/accentBorderCallout の1〜3)は、
            // 引き出し線の頂点を (adj1=y1, adj2=x1), (adj3=y2, adj4=x2), ... の組で持つ(要件10.14)。
            // 折れ数Nの吹き出しは N+1 個の頂点 = 2(N+1) 個のガイドを持つ。
            [ShapePresetType.Callout1] = LineCalloutGuideNames1,
            [ShapePresetType.Callout2] = LineCalloutGuideNames2,
            [ShapePresetType.Callout3] = LineCalloutGuideNames3,
            [ShapePresetType.BorderCallout1] = LineCalloutGuideNames1,
            [ShapePresetType.BorderCallout2] = LineCalloutGuideNames2,
            [ShapePresetType.BorderCallout3] = LineCalloutGuideNames3,
            [ShapePresetType.AccentCallout1] = LineCalloutGuideNames1,
            [ShapePresetType.AccentCallout2] = LineCalloutGuideNames2,
            [ShapePresetType.AccentCallout3] = LineCalloutGuideNames3,
            [ShapePresetType.AccentBorderCallout1] = LineCalloutGuideNames1,
            [ShapePresetType.AccentBorderCallout2] = LineCalloutGuideNames2,
            [ShapePresetType.AccentBorderCallout3] = LineCalloutGuideNames3,
            // star4/5/6/8はECMA-376既定で単一の調整ガイド"adj"(内側頂点の半径比)を持つ。
            [ShapePresetType.Star4] = new[] { "adj" },
            [ShapePresetType.Star5] = new[] { "adj" },
            [ShapePresetType.Star6] = new[] { "adj" },
            [ShapePresetType.Star8] = new[] { "adj" },
            // フローチャート記号は本プロダクトでは固定比率の形状として描画し、調整ガイド値は読み取らない(design.md参照)。
            [ShapePresetType.FlowChartProcess] = Array.Empty<string>(),
            [ShapePresetType.FlowChartDecision] = Array.Empty<string>(),
            [ShapePresetType.FlowChartTerminator] = Array.Empty<string>(),
            [ShapePresetType.FlowChartInputOutput] = Array.Empty<string>(),
            [ShapePresetType.FlowChartDocument] = Array.Empty<string>(),
            [ShapePresetType.FlowChartPredefinedProcess] = Array.Empty<string>(),
            [ShapePresetType.FlowChartConnector] = Array.Empty<string>(),
        };

        /// <summary>
        /// 1シートに含める図形・接続線・グループの合計数(グループ内の子孫を含む)の上限。
        /// 画像の上限(<see cref="MaxImagesPerSheet"/>)とは独立にカウントする
        /// (security-reviewer指摘と同じ考え方。要件10.8)。
        /// </summary>
        internal const int MaxShapesPerSheet = 50;

        /// <summary>
        /// 図形1つに含まれる全テキスト(段落・ランを連結した文字数)の上限。極端に長い文字列に
        /// 対する折り返し計算量を避けるための安全弁(要件10.8)。
        /// </summary>
        private const int MaxShapeTextLength = 2000;

        /// <summary>
        /// グループ(<c>xdr:grpSp</c>)のネスト段数の上限。極端に深いネストによる再帰処理の
        /// 計算量を避けるための安全弁(要件10.8)。トップレベルのグループ自身を1段目とする。
        /// </summary>
        internal const int MaxShapeNestingDepth = 5;

        /// <summary>
        /// 1つの塗りつぶし(<c>a:gsLst</c>)に含める、グラデーションストップ(<c>a:gs</c>)の
        /// 個数の上限。Excel UI自体が通常扱う範囲(数個〜十数個)を大きく超える値とし、
        /// 実用上の妨げにはならない範囲で、大量の<c>a:gs</c>要素を仕込んだ入力による
        /// メモリ・CPU消費を抑える(要件10.8。security-reviewer指摘)。
        /// </summary>
        private const int MaxGradientStopsPerFill = 64;

        /// <summary>
        /// シートに埋め込まれた画像(<c>xdr:pic</c>)・図形(<c>xdr:sp</c>)・接続線(<c>xdr:cxnSp</c>)・
        /// グループ(<c>xdr:grpSp</c>)を読み取る(要件9, 10)。
        /// </summary>
        /// <remarks>
        /// <c>drawing.xml</c> のアンカーを出現順に1回だけ列挙し、これらが混在する場合の
        /// 重なり順(z-order。要件10.3)を保った1つのリストを返す。これら4種以外の描画
        /// オブジェクト(図表枠`xdr:graphicFrame`・グラフ等)は対象外とし、
        /// <see cref="DetectUnsupportedElements"/>側で引き続き「サポート外要素」として扱う
        /// (要件9.5, 10.7)。画像・対応済みプリセットの図形/接続線、グループは帳票定義の
        /// <c>unsupportedElements</c> 設定によらず常に読み取り対象とするが、デコード不能な
        /// 画像形式・非対応プリセット・上限を超える枚数/サイズ/文字数/ネスト段数だけは
        /// 同じ設定に従う(要件9.4, 9.6, 10.7, 10.8)。
        /// </remarks>
        private static List<DrawingObjectModel> ReadDrawingObjects(
            string sheetName, WorksheetPart worksheetPart, DrawingColorResolver colors, WorkbookReadOptions options)
        {
            var result = new List<DrawingObjectModel>();
            var drawing = worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawing is null)
            {
                return result;
            }

            var drawingsPart = worksheetPart.DrawingsPart!;
            var imageCount = 0;
            var shapeCount = 0;

            foreach (var anchor in drawing.ChildElements)
            {
                var fromMarker = anchor switch
                {
                    Xdr.TwoCellAnchor two => two.FromMarker,
                    Xdr.OneCellAnchor one => one.FromMarker,
                    _ => null,
                };

                if (fromMarker is null || !TryReadMarker(fromMarker, out var anchorCell, out var anchorOffset))
                {
                    continue;
                }

                var picture = anchor switch
                {
                    Xdr.TwoCellAnchor two => two.GetFirstChild<Xdr.Picture>(),
                    Xdr.OneCellAnchor one => one.GetFirstChild<Xdr.Picture>(),
                    _ => null,
                };

                if (picture is not null)
                {
                    var image = ReadImage(
                        sheetName, drawingsPart, anchor, picture, anchorCell, anchorOffset, options, ref imageCount);
                    if (image is not null)
                    {
                        result.Add(image);
                    }

                    continue;
                }

                var shape = anchor switch
                {
                    Xdr.TwoCellAnchor two => two.GetFirstChild<Xdr.Shape>(),
                    Xdr.OneCellAnchor one => one.GetFirstChild<Xdr.Shape>(),
                    _ => null,
                };

                if (shape is not null)
                {
                    var shapeModel = ReadShape(sheetName, anchor, shape, anchorCell, anchorOffset, colors, options, ref shapeCount);
                    if (shapeModel is not null)
                    {
                        result.Add(shapeModel);
                    }

                    continue;
                }

                var connector = anchor switch
                {
                    Xdr.TwoCellAnchor two => two.GetFirstChild<Xdr.ConnectionShape>(),
                    Xdr.OneCellAnchor one => one.GetFirstChild<Xdr.ConnectionShape>(),
                    _ => null,
                };

                if (connector is not null)
                {
                    var connectorModel = ReadConnector(sheetName, anchor, connector, anchorCell, anchorOffset, colors, options, ref shapeCount);
                    if (connectorModel is not null)
                    {
                        result.Add(connectorModel);
                    }

                    continue;
                }

                var group = anchor switch
                {
                    Xdr.TwoCellAnchor two => two.GetFirstChild<Xdr.GroupShape>(),
                    Xdr.OneCellAnchor one => one.GetFirstChild<Xdr.GroupShape>(),
                    _ => null,
                };

                if (group is not null)
                {
                    var groupModel = ReadGroupShape(
                        sheetName, drawingsPart, anchor, group, anchorCell, anchorOffset, colors, options, ref shapeCount, ref imageCount);
                    if (groupModel is not null)
                    {
                        result.Add(groupModel);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// アンカー(<see cref="Xdr.TwoCellAnchor"/>/<see cref="Xdr.OneCellAnchor"/>)から
        /// 画像・図形共通の終端(サイズ)を読み取る。
        /// </summary>
        private static AnchorExtent? ReadAnchorExtent(OpenXmlElement anchor) => anchor switch
        {
            Xdr.TwoCellAnchor two when two.ToMarker is { } toMarker && TryReadMarker(toMarker, out var toCell, out var toOffset) =>
                new CellSpanAnchorExtent(toCell, toOffset),
            Xdr.OneCellAnchor { Extent: { Cx: { } cx, Cy: { } cy } } =>
                new FixedAnchorExtent(Units.EmusToPoints(cx.Value), Units.EmusToPoints(cy.Value)),
            _ => null,
        };

        /// <summary>1つの<c>xdr:pic</c>アンカーを<see cref="ImageModel"/>として読み取る(要件9)。</summary>
        private static ImageModel? ReadImage(
            string sheetName,
            DrawingsPart drawingsPart,
            OpenXmlElement anchor,
            Xdr.Picture picture,
            CellAddress anchorCell,
            PointPt anchorOffset,
            WorkbookReadOptions options,
            ref int imageCount)
        {
            if (imageCount >= MaxImagesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の画像の数が上限({MaxImagesPerSheet}枚)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyImages",
                        options.ReportCode,
                        sheetName);
                }

                // ignore時はこのアンカーだけを無視し、以降のアンカーの走査は続ける
                // (画像・図形が混在するシートで、他方の走査を止めないため)。
                return null;
            }

            // 読み取りを試みた時点で数える。検証に失敗した画像を数えないと、同じ画像パートを参照する
            // 壊れたアンカーを大量に並べて、上限に達しないまま画像を何度も展開させられる(security-reviewer指摘)。
            imageCount++;

            // ReadAnchorExtentより先に検証する。壊れたアンカー(ReadAnchorExtentの失敗)は
            // モードによらず常に無言でスキップするため、先に判定すると不正な画像形式/
            // サイズ超過/シグネチャ不一致がErrorモードでも例外化されずに握りつぶされてしまう
            // (code-reviewer指摘。リファクタ前の検証順序を維持する)。
            if (!TryReadValidatedImage(sheetName, drawingsPart, picture, options, out var data, out var contentType))
            {
                return null;
            }

            var extent = ReadAnchorExtent(anchor);
            if (extent is null)
            {
                return null;
            }

            var id = ReadShapeId(picture.NonVisualPictureProperties?.NonVisualDrawingProperties);
            var rotationDegrees = (picture.ShapeProperties?.Transform2D?.Rotation?.Value ?? 0) / 60000.0;

            return new ImageModel(id, data, contentType, rotationDegrees, anchorCell, anchorOffset, extent);
        }

        /// <summary>
        /// <c>xdr:pic</c>(トップレベル・グループ内共通)から画像バイナリを読み取り、
        /// content-type許可リスト・サイズ上限・シグネチャの検証を行う(要件9.4, 9.6)。
        /// 検証に失敗した場合、Errorモードなら例外を送出し、Ignoreモードなら<c>false</c>を返す。
        /// </summary>
        private static bool TryReadValidatedImage(
            string sheetName,
            DrawingsPart drawingsPart,
            Xdr.Picture picture,
            WorkbookReadOptions options,
            out byte[] data,
            out string contentType)
        {
            data = Array.Empty<byte>();
            contentType = string.Empty;

            var embedId = picture.BlipFill?.Blip?.Embed?.Value;
            if (string.IsNullOrEmpty(embedId))
            {
                return false;
            }

            OpenXmlPart embeddedPart;
            try
            {
                embeddedPart = drawingsPart.GetPartById(embedId!);
            }
            catch (ArgumentOutOfRangeException)
            {
                // 参照先のパートがファイル内に無い(要件6.10)。SDKの例外を漏らさず、サポート外要素として扱う。
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の画像の参照先('{embedId}')がファイル内にありません。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "MissingImagePart",
                        options.ReportCode,
                        sheetName);
                }

                return false;
            }

            if (embeddedPart is not ImagePart imagePart)
            {
                return false;
            }

            if (!SupportedImageContentTypes.Contains(imagePart.ContentType))
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の画像形式 '{imagePart.ContentType}' には対応していません"
                        + "(対応形式: PNG/JPEG/GIF/BMP)。帳票定義の unsupportedElements が 'error' のため中止します。",
                        "UnsupportedImageFormat",
                        options.ReportCode,
                        sheetName);
                }

                return false;
            }

            using (var stream = imagePart.GetStream())
            {
                if (!TryReadBounded(stream, MaxImageDataBytes, out data))
                {
                    if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                    {
                        throw new UnsupportedWorkbookElementException(
                            $"シート '{sheetName}' の画像のサイズが上限({MaxImageDataBytes / (1024 * 1024)}MB)を"
                            + "超えています。帳票定義の unsupportedElements が 'error' のため中止します。",
                            "ImageTooLarge",
                            options.ReportCode,
                            sheetName);
                    }

                    return false;
                }
            }

            // ContentTypeはOPCパッケージ側の申告値に過ぎず、実際のバイト列と一致する保証がない。
            // ネイティブコードのデコーダ(SkiaSharp)に渡す前に、ファイル先頭のシグネチャで
            // 最低限の裏取りを行う(security-reviewer指摘)。
            if (!MatchesContentTypeSignature(imagePart.ContentType, data))
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の画像データが '{imagePart.ContentType}' として不正です"
                        + "(ファイル先頭のシグネチャが一致しません)。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "UnsupportedImageFormat",
                        options.ReportCode,
                        sheetName);
                }

                return false;
            }

            contentType = imagePart.ContentType;
            return true;
        }

        /// <summary>1つの<c>xdr:sp</c>アンカーを<see cref="ShapeModel"/>として読み取る(要件10)。</summary>
        private static ShapeModel? ReadShape(
            string sheetName,
            OpenXmlElement anchor,
            Xdr.Shape shape,
            CellAddress anchorCell,
            PointPt anchorOffset,
            DrawingColorResolver colors,
            WorkbookReadOptions options,
            ref int shapeCount)
        {
            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                // ignore時はこのアンカーだけを無視し、以降のアンカーの走査は続ける。
                return null;
            }

            var shapeProperties = shape.ShapeProperties;
            var presetGeometry = shapeProperties?.GetFirstChild<Dr.PresetGeometry>();
            var presetValue = presetGeometry?.Preset?.Value;

            if (presetValue is null || !SupportedShapePresets.TryGetValue(presetValue.Value, out var preset))
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    var presetDescription = presetValue is null ? "(prstGeomなし/custGeom)" : presetValue.Value.ToString();
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形のプリセットジオメトリ '{presetDescription}' には対応していません。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "UnsupportedShapePreset",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var extent = ReadAnchorExtent(anchor);
            if (extent is null)
            {
                return null;
            }

            var adjustmentValues = ReadShapeAdjustmentValues(preset, presetGeometry);
            var rotationDegrees = (shapeProperties?.Transform2D?.Rotation?.Value ?? 0) / 60000.0;
            var fill = ReadShapeFill(shapeProperties, shape.ShapeStyle, colors);
            var outline = ReadShapeOutline(shapeProperties, shape.ShapeStyle, colors);
            var text = ReadShapeText(shape.TextBody, shape.ShapeStyle, colors, out var textTooLong);

            if (textTooLong)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形のテキストが上限({MaxShapeTextLength}文字)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "ShapeTextTooLong",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var id = ReadShapeId(shape.NonVisualShapeProperties?.NonVisualDrawingProperties);

            shapeCount++;
            return new ShapeModel(
                id, preset, adjustmentValues, rotationDegrees, fill, outline, text, anchorCell, anchorOffset, extent,
                shapeProperties?.Transform2D?.HorizontalFlip?.Value ?? false,
                shapeProperties?.Transform2D?.VerticalFlip?.Value ?? false);
        }

        /// <summary>1つの<c>xdr:cxnSp</c>アンカーを<see cref="ConnectorModel"/>として読み取る(要件10.9)。</summary>
        private static ConnectorModel? ReadConnector(
            string sheetName,
            OpenXmlElement anchor,
            Xdr.ConnectionShape connector,
            CellAddress anchorCell,
            PointPt anchorOffset,
            DrawingColorResolver colors,
            WorkbookReadOptions options,
            ref int shapeCount)
        {
            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var shapeProperties = connector.ShapeProperties;
            var presetGeometry = shapeProperties?.GetFirstChild<Dr.PresetGeometry>();
            var presetValue = presetGeometry?.Preset?.Value;

            if (presetValue is null || !SupportedConnectorPresets.TryGetValue(presetValue.Value, out var preset))
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    var presetDescription = presetValue is null ? "(prstGeomなし/custGeom)" : presetValue.Value.ToString();
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の接続線のプリセットジオメトリ '{presetDescription}' には対応していません。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "UnsupportedShapePreset",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var extent = ReadAnchorExtent(anchor);
            if (extent is null)
            {
                return null;
            }

            var transform = shapeProperties?.Transform2D;
            var rotationDegrees = (transform?.Rotation?.Value ?? 0) / 60000.0;
            var flipHorizontal = transform?.HorizontalFlip?.Value ?? false;
            var flipVertical = transform?.VerticalFlip?.Value ?? false;
            var outline = ReadConnectorOutline(shapeProperties, connector.ShapeStyle, colors);

            var connectorShapeDrawingProperties =
                connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
            var startConnection = ReadConnectionRef(connectorShapeDrawingProperties?.StartConnection);
            var endConnection = ReadConnectionRef(connectorShapeDrawingProperties?.EndConnection);

            shapeCount++;
            return new ConnectorModel(
                preset, rotationDegrees, flipHorizontal, flipVertical, outline,
                startConnection, endConnection, anchorCell, anchorOffset, extent);
        }

        /// <summary>1つの<c>xdr:grpSp</c>アンカーを<see cref="GroupShapeModel"/>として読み取る(要件10.10)。</summary>
        private static GroupShapeModel? ReadGroupShape(
            string sheetName,
            DrawingsPart drawingsPart,
            OpenXmlElement anchor,
            Xdr.GroupShape group,
            CellAddress anchorCell,
            PointPt anchorOffset,
            DrawingColorResolver colors,
            WorkbookReadOptions options,
            ref int shapeCount,
            ref int imageCount)
        {
            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var extent = ReadAnchorExtent(anchor);
            if (extent is null)
            {
                return null;
            }

            var transformGroup = group.GroupShapeProperties?.TransformGroup;
            var childOffset = ReadChildOffset(transformGroup?.ChildOffset);
            var childExtent = ReadChildExtent(transformGroup?.ChildExtents);
            if (childOffset is null || childExtent is null)
            {
                return null;
            }

            var rotationDegrees = (transformGroup?.Rotation?.Value ?? 0) / 60000.0;

            var children = ReadGroupChildren(sheetName, drawingsPart, group, colors, options, ref shapeCount, ref imageCount, currentGroupDepth: 1);
            if (children is null)
            {
                // Ignoreモードで内部に非対応要素があった場合。グループの一部だけを描画すると
                // 意図しない見た目になるため、グループ全体を破棄する(要件10.10)。
                return null;
            }

            // 子孫の読み取りでshapeCountが増加しているため、グループ自身を1件として
            // 数える前に改めて上限を確認する(code-reviewer指摘。子孫読み取り前のチェックだけでは
            // グループ自身の分の加算で上限をわずかに超過しうる)。
            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var id = ReadShapeId(group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties);

            shapeCount++;
            return new GroupShapeModel(
                id, childOffset.Value, childExtent.Value, children, rotationDegrees, anchorCell, anchorOffset, extent,
                transformGroup?.HorizontalFlip?.Value ?? false,
                transformGroup?.VerticalFlip?.Value ?? false);
        }

        /// <summary>
        /// グループ(<see cref="Xdr.GroupShape"/>)直下の子要素(<c>xdr:sp</c>/<c>xdr:pic</c>/
        /// <c>xdr:cxnSp</c>/入れ子の<c>xdr:grpSp</c>)を出現順に<see cref="GroupChildModel"/>へ
        /// 変換する(要件10.10)。子孫のいずれか1つでも非対応(非対応プリセット・非対応の
        /// 描画オブジェクト種別・上限超過・ネスト過多)であれば、Errorモードは即座に例外を送出し、
        /// Ignoreモードは<c>null</c>を返してグループ全体を呼び出し元に破棄させる。
        /// </summary>
        private static IReadOnlyList<GroupChildModel>? ReadGroupChildren(
            string sheetName,
            DrawingsPart drawingsPart,
            Xdr.GroupShape group,
            DrawingColorResolver colors,
            WorkbookReadOptions options,
            ref int shapeCount,
            ref int imageCount,
            int currentGroupDepth)
        {
            var children = new List<GroupChildModel>();

            foreach (var element in group.ChildElements)
            {
                if (element is Xdr.NonVisualGroupShapeProperties or Xdr.GroupShapeProperties)
                {
                    // グループ自身のメタデータ(グループ全体の変形情報等)であり、描画対象の
                    // 子要素ではないため走査対象外とする。
                    continue;
                }

                var isSupportedChildType = element is Xdr.Shape or Xdr.Picture or Xdr.ConnectionShape or Xdr.GroupShape;
                GroupChildModel? child = element switch
                {
                    Xdr.Shape shape => ReadGroupChildShape(sheetName, shape, colors, options, ref shapeCount),
                    Xdr.Picture picture => ReadGroupChildImage(sheetName, drawingsPart, picture, options, ref imageCount),
                    Xdr.ConnectionShape connector => ReadGroupChildConnector(sheetName, connector, colors, options, ref shapeCount),
                    Xdr.GroupShape nestedGroup => ReadGroupChildGroup(
                        sheetName, drawingsPart, nestedGroup, colors, options, ref shapeCount, ref imageCount, currentGroupDepth + 1),
                    _ => null,
                };

                if (child is null)
                {
                    if (!isSupportedChildType && options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                    {
                        throw new UnsupportedWorkbookElementException(
                            $"シート '{sheetName}' のグループ内に非対応の描画オブジェクト(図表枠等)が含まれています。"
                            + "帳票定義の unsupportedElements が 'error' のため中止します。",
                            "UnsupportedShapePreset",
                            options.ReportCode,
                            sheetName);
                    }

                    // 個別の子要素側でErrorモードならすでに例外を送出済み。ここに到達するのは
                    // Ignoreモードでの非対応(非対応の型/プリセット/上限超過等)であり、グループ全体を破棄する。
                    return null;
                }

                children.Add(child);
            }

            return children;
        }

        /// <summary>グループ内の<c>xdr:sp</c>子要素を<see cref="GroupChildShape"/>として読み取る(要件10.10)。</summary>
        private static GroupChildShape? ReadGroupChildShape(
            string sheetName, Xdr.Shape shape, DrawingColorResolver colors, WorkbookReadOptions options, ref int shapeCount)
        {
            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var shapeProperties = shape.ShapeProperties;
            var presetGeometry = shapeProperties?.GetFirstChild<Dr.PresetGeometry>();
            var presetValue = presetGeometry?.Preset?.Value;

            if (presetValue is null || !SupportedShapePresets.TryGetValue(presetValue.Value, out var preset))
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    var presetDescription = presetValue is null ? "(prstGeomなし/custGeom)" : presetValue.Value.ToString();
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' のグループ内図形のプリセットジオメトリ '{presetDescription}' には対応していません。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "UnsupportedShapePreset",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var localRect = ReadLocalRect(shapeProperties?.Transform2D);
            if (localRect is null)
            {
                return null;
            }

            var adjustmentValues = ReadShapeAdjustmentValues(preset, presetGeometry);
            var rotationDegrees = (shapeProperties?.Transform2D?.Rotation?.Value ?? 0) / 60000.0;
            var fill = ReadShapeFill(shapeProperties, shape.ShapeStyle, colors);
            var outline = ReadShapeOutline(shapeProperties, shape.ShapeStyle, colors);
            var text = ReadShapeText(shape.TextBody, shape.ShapeStyle, colors, out var textTooLong);

            if (textTooLong)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' のグループ内図形のテキストが上限({MaxShapeTextLength}文字)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "ShapeTextTooLong",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var id = ReadShapeId(shape.NonVisualShapeProperties?.NonVisualDrawingProperties);

            shapeCount++;
            return new GroupChildShape(
                id, localRect.Value, preset, adjustmentValues, rotationDegrees, fill, outline, text,
                shapeProperties?.Transform2D?.HorizontalFlip?.Value ?? false,
                shapeProperties?.Transform2D?.VerticalFlip?.Value ?? false);
        }

        /// <summary>グループ内の<c>xdr:pic</c>子要素を<see cref="GroupChildImage"/>として読み取る(要件10.10)。</summary>
        private static GroupChildImage? ReadGroupChildImage(
            string sheetName, DrawingsPart drawingsPart, Xdr.Picture picture, WorkbookReadOptions options, ref int imageCount)
        {
            if (imageCount >= MaxImagesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の画像の数が上限({MaxImagesPerSheet}枚)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyImages",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            // トップレベルの画像と同じく、読み取りを試みた時点で数える。
            imageCount++;

            if (!TryReadValidatedImage(sheetName, drawingsPart, picture, options, out var data, out var contentType))
            {
                return null;
            }

            var localRect = ReadLocalRect(picture.ShapeProperties?.Transform2D);
            if (localRect is null)
            {
                return null;
            }

            var id = ReadShapeId(picture.NonVisualPictureProperties?.NonVisualDrawingProperties);
            var rotationDegrees = (picture.ShapeProperties?.Transform2D?.Rotation?.Value ?? 0) / 60000.0;

            return new GroupChildImage(id, localRect.Value, data, contentType, rotationDegrees);
        }

        /// <summary>グループ内の<c>xdr:cxnSp</c>子要素を<see cref="GroupChildConnector"/>として読み取る(要件10.10)。</summary>
        private static GroupChildConnector? ReadGroupChildConnector(
            string sheetName, Xdr.ConnectionShape connector, DrawingColorResolver colors, WorkbookReadOptions options, ref int shapeCount)
        {
            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var shapeProperties = connector.ShapeProperties;
            var presetGeometry = shapeProperties?.GetFirstChild<Dr.PresetGeometry>();
            var presetValue = presetGeometry?.Preset?.Value;

            if (presetValue is null || !SupportedConnectorPresets.TryGetValue(presetValue.Value, out var preset))
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    var presetDescription = presetValue is null ? "(prstGeomなし/custGeom)" : presetValue.Value.ToString();
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' のグループ内接続線のプリセットジオメトリ '{presetDescription}' には対応していません。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "UnsupportedShapePreset",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var localRect = ReadLocalRect(shapeProperties?.Transform2D);
            if (localRect is null)
            {
                return null;
            }

            var transform = shapeProperties?.Transform2D;
            var rotationDegrees = (transform?.Rotation?.Value ?? 0) / 60000.0;
            var flipHorizontal = transform?.HorizontalFlip?.Value ?? false;
            var flipVertical = transform?.VerticalFlip?.Value ?? false;
            var outline = ReadConnectorOutline(shapeProperties, connector.ShapeStyle, colors);

            var connectorShapeDrawingProperties =
                connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
            var startConnection = ReadConnectionRef(connectorShapeDrawingProperties?.StartConnection);
            var endConnection = ReadConnectionRef(connectorShapeDrawingProperties?.EndConnection);

            shapeCount++;
            return new GroupChildConnector(
                localRect.Value, preset, rotationDegrees, flipHorizontal, flipVertical, outline,
                startConnection, endConnection);
        }

        /// <summary>
        /// グループ内の入れ子の<c>xdr:grpSp</c>子要素を<see cref="GroupChildGroup"/>として
        /// 読み取る(要件10.10)。<paramref name="depth"/>が<see cref="MaxShapeNestingDepth"/>を
        /// 超える場合は拒否する(要件10.8)。
        /// </summary>
        private static GroupChildGroup? ReadGroupChildGroup(
            string sheetName,
            DrawingsPart drawingsPart,
            Xdr.GroupShape group,
            DrawingColorResolver colors,
            WorkbookReadOptions options,
            ref int shapeCount,
            ref int imageCount,
            int depth)
        {
            if (depth > MaxShapeNestingDepth)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' のグループのネストが上限({MaxShapeNestingDepth}段)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "GroupNestingTooDeep",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を超えています。"
                        + "帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var transformGroup = group.GroupShapeProperties?.TransformGroup;
            var localRect = ReadLocalRect(transformGroup);
            var childOffset = ReadChildOffset(transformGroup?.ChildOffset);
            var childExtent = ReadChildExtent(transformGroup?.ChildExtents);
            if (localRect is null || childOffset is null || childExtent is null)
            {
                return null;
            }

            var rotationDegrees = (transformGroup?.Rotation?.Value ?? 0) / 60000.0;

            var children = ReadGroupChildren(sheetName, drawingsPart, group, colors, options, ref shapeCount, ref imageCount, depth);
            if (children is null)
            {
                return null;
            }

            // 子孫の読み取りでshapeCountが増加しているため、グループ自身を1件として
            // 数える前に改めて上限を確認する(code-reviewer指摘)。
            if (shapeCount >= MaxShapesPerSheet)
            {
                if (options.UnsupportedElementBehavior == UnsupportedElementBehavior.Error)
                {
                    throw new UnsupportedWorkbookElementException(
                        $"シート '{sheetName}' の図形・接続線・グループの数が上限({MaxShapesPerSheet}個)を"
                        + "超えています。帳票定義の unsupportedElements が 'error' のため中止します。",
                        "TooManyShapes",
                        options.ReportCode,
                        sheetName);
                }

                return null;
            }

            var id = ReadShapeId(group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties);

            shapeCount++;
            return new GroupChildGroup(
                id, localRect.Value, rotationDegrees, childOffset.Value, childExtent.Value, children,
                transformGroup?.HorizontalFlip?.Value ?? false,
                transformGroup?.VerticalFlip?.Value ?? false);
        }

        /// <summary>
        /// <c>NonVisualDrawingProperties/@id</c>(要件10.11)を読み取る。接続線の接続先解決の
        /// キーとしてのみ使うため、要素が無い(理論上は起こらない)場合は既定値0とする。
        /// </summary>
        private static uint ReadShapeId(Xdr.NonVisualDrawingProperties? nonVisualDrawingProperties) =>
            nonVisualDrawingProperties?.Id?.Value ?? 0;

        /// <summary>
        /// <c>a:stCxn</c>/<c>a:endCxn</c>(要件10.11)を<see cref="ConnectionRef"/>として読み取る。
        /// 要素が無ければ<c>null</c>(接続線の始点/終点がどの図形にも紐づいていない)。
        /// </summary>
        private static ConnectionRef? ReadConnectionRef(Dr.ConnectionType? connection) =>
            connection is { Id: { } id, Index: { } index } ? new ConnectionRef(id.Value, index.Value) : null;

        /// <summary>グループの子座標空間の原点(<c>a:chOff</c>)をポイント単位で読み取る。</summary>
        private static PointPt? ReadChildOffset(Dr.ChildOffset? childOffset) =>
            childOffset is { X: { } x, Y: { } y } ? new PointPt(Units.EmusToPoints(x.Value), Units.EmusToPoints(y.Value)) : (PointPt?)null;

        /// <summary>グループの子座標空間の大きさ(<c>a:chExt</c>)をポイント単位で読み取る。</summary>
        private static PointPt? ReadChildExtent(Dr.ChildExtents? childExtents) =>
            childExtents is { Cx: { } cx, Cy: { } cy }
                ? new PointPt(Units.EmusToPoints(cx.Value), Units.EmusToPoints(cy.Value))
                : (PointPt?)null;

        /// <summary>グループ内要素の位置・サイズ(<c>a:off</c>/<c>a:ext</c>)を親の子座標空間上の矩形として読み取る。</summary>
        private static RectPt? ReadLocalRect(Dr.Transform2D? transform) =>
            transform is { Offset: { X: { } x, Y: { } y }, Extents: { Cx: { } cx, Cy: { } cy } } && cx.Value >= 0 && cy.Value >= 0
                ? new RectPt(Units.EmusToPoints(x.Value), Units.EmusToPoints(y.Value), Units.EmusToPoints(cx.Value), Units.EmusToPoints(cy.Value))
                : (RectPt?)null;

        /// <summary>入れ子グループ自身の位置・サイズ(<c>a:off</c>/<c>a:ext</c>)を親の子座標空間上の矩形として読み取る。</summary>
        private static RectPt? ReadLocalRect(Dr.TransformGroup? transform) =>
            transform is { Offset: { X: { } x, Y: { } y }, Extents: { Cx: { } cx, Cy: { } cy } } && cx.Value >= 0 && cy.Value >= 0
                ? new RectPt(Units.EmusToPoints(x.Value), Units.EmusToPoints(y.Value), Units.EmusToPoints(cx.Value), Units.EmusToPoints(cy.Value))
                : (RectPt?)null;

        /// <summary>
        /// プリセットごとの調整ガイド(<see cref="ShapeAdjustmentGuideNames"/>)の並び順に対応する
        /// 固定長のリストを組み立てる。ファイルに該当ガイドが無い位置は <see cref="double.NaN"/> とし、
        /// Renderingレイヤーがその位置のECMA-376既定値を補う(design.md「Parsing レイヤー」参照)。
        /// </summary>
        private static IReadOnlyList<double> ReadShapeAdjustmentValues(ShapePresetType preset, Dr.PresetGeometry? presetGeometry)
        {
            var guideNames = ShapeAdjustmentGuideNames[preset];
            if (guideNames.Length == 0)
            {
                return Array.Empty<double>();
            }

            var guideValues = new Dictionary<string, double>(StringComparer.Ordinal);
            if (presetGeometry?.AdjustValueList is { } adjustValueList)
            {
                foreach (var guide in adjustValueList.Elements<Dr.ShapeGuide>())
                {
                    if (guide.Name?.Value is { } name && TryParseGuideFormula(guide.Formula?.Value, out var value))
                    {
                        guideValues[name] = value;
                    }
                }
            }

            var result = new double[guideNames.Length];
            for (var i = 0; i < guideNames.Length; i++)
            {
                result[i] = guideValues.TryGetValue(guideNames[i], out var value) ? value : double.NaN;
            }

            return result;
        }

        /// <summary><c>a:gd/@fmla</c>(例: <c>"val 16667"</c>)から比率(0〜1)を読み取る。</summary>
        private static bool TryParseGuideFormula(string? formula, out double value)
        {
            value = default;
            if (string.IsNullOrEmpty(formula))
            {
                return false;
            }

            var parts = formula!.Split(' ');
            var numberText = parts[parts.Length - 1];
            if (!int.TryParse(numberText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var raw))
            {
                return false;
            }

            value = raw / 100000.0;
            return true;
        }

        /// <summary>
        /// 図形の塗りつぶし(<c>a:noFill</c>/<c>a:solidFill</c>/<c>a:gradFill</c>)を読み取る。<c>spPr</c>に塗りつぶしの
        /// 指定が無ければ、図形のスタイルの<c>fillRef</c>から求める(要件10.15, 10.16)。
        /// </summary>
        private static ShapeFill? ReadShapeFill(
            Xdr.ShapeProperties? shapeProperties, Xdr.ShapeStyle? style, DrawingColorResolver colors)
        {
            if (shapeProperties?.GetFirstChild<Dr.NoFill>() is not null)
            {
                return null;
            }

            if (shapeProperties?.GetFirstChild<Dr.GradientFill>() is { } gradientFill)
            {
                var stops = ReadGradientStops(gradientFill.GradientStopList, colors);
                if (stops is null)
                {
                    return null;
                }

                if (gradientFill.GetFirstChild<Dr.PathGradientFill>() is { } pathFill)
                {
                    // a:path[@path='circle'/'rect'/'shape']のいずれも放射状として近似する
                    // (要件10.6補足)。中心はa:fillToRectの中心、無ければ矩形中心(0.5, 0.5)。
                    var toRect = pathFill.FillToRectangle;
                    var center = new PointPt(
                        (FillToRectFraction(toRect?.Left) + FillToRectFraction(toRect?.Right)) / 2.0,
                        (FillToRectFraction(toRect?.Top) + FillToRectFraction(toRect?.Bottom)) / 2.0);
                    return new RadialGradientShapeFill(stops, center);
                }

                var angle = gradientFill.GetFirstChild<Dr.LinearGradientFill>()?.Angle?.Value ?? 0;
                return new LinearGradientShapeFill(stops, angle / 60000.0);
            }

            if (shapeProperties?.GetFirstChild<Dr.SolidFill>() is { } solidFill)
            {
                return colors.TryResolve(solidFill, placeholder: null, out var color) ? new SolidShapeFill(color) : null;
            }

            // 画像・模様・グループの塗りつぶしは対応していない。スタイルの塗りつぶしで代用しない。
            if (shapeProperties?.GetFirstChild<Dr.BlipFill>() is not null
                || shapeProperties?.GetFirstChild<Dr.PatternFill>() is not null
                || shapeProperties?.GetFirstChild<Dr.GroupFill>() is not null)
            {
                return null;
            }

            // fillRef/@idx が 0 なら塗りなし。1以上はテーマの塗りつぶしの書式を指すが、Excelの既定テーマでは
            // いずれもスタイルの色(phClr)を基にした塗りのため、その色の単色で近似する(要件10.16)。
            if (style?.FillReference is { } fillRef
                && fillRef.Index?.Value is { } fillIndex && fillIndex > 0
                && colors.TryResolve(fillRef, placeholder: null, out var styleColor))
            {
                return new SolidShapeFill(styleColor);
            }

            return null;
        }

        /// <summary>
        /// <c>a:gsLst</c> の全ストップ(位置・色)を読み取る(要件10.6。3点以上に対応)。
        /// 位置・色のいずれかを読み取れないストップが1つでもあれば全体を<c>null</c>とし、
        /// 呼び出し側で塗りなしにフォールバックする(画像対応時の「不正な入力は無視する」方針と同様)。
        /// </summary>
        private static IReadOnlyList<GradientStop>? ReadGradientStops(Dr.GradientStopList? gradientStopList, DrawingColorResolver colors)
        {
            if (gradientStopList is null)
            {
                return null;
            }

            var stops = new List<GradientStop>();
            foreach (var stop in gradientStopList.Elements<Dr.GradientStop>())
            {
                // 件数を数える前に上限で打ち切る(shapeCount/imageCount/MaxShapeTextLengthと
                // 同じ「無制限の入力から先に上限チェックする」方針。security-reviewer指摘)。
                if (stops.Count >= MaxGradientStopsPerFill)
                {
                    return null;
                }

                if (stop.Position?.Value is not { } position || !colors.TryResolve(stop, placeholder: null, out var color))
                {
                    return null;
                }

                stops.Add(new GradientStop(PermilleToFraction(position), color));
            }

            return stops.Count >= 2 ? stops : null;
        }

        /// <summary>OOXMLの千分率(0〜100000)を0.0〜1.0の比率に変換する。</summary>
        private static double PermilleToFraction(int value) => value / 100000.0;

        /// <summary><c>a:fillToRect</c>の1辺(<see cref="Int32Value"/>)を比率に変換する。無指定なら50%とする。</summary>
        private static double FillToRectFraction(Int32Value? value) => value?.Value is { } v ? PermilleToFraction(v) : 0.5;

        /// <summary>
        /// 図形・接続線の枠線(<c>a:ln</c>)を読み取る。<c>a:ln</c>に色の指定が無ければ、図形のスタイルの
        /// <c>lnRef</c>から色と太さを求める(要件10.15, 10.16)。
        /// </summary>
        private static ShapeOutline? ReadShapeOutline(
            Xdr.ShapeProperties? shapeProperties, Xdr.ShapeStyle? style, DrawingColorResolver colors)
        {
            var outline = shapeProperties?.GetFirstChild<Dr.Outline>();
            if (outline?.GetFirstChild<Dr.NoFill>() is not null)
            {
                return null;
            }

            var lineRef = style?.LineReference is { } reference && reference.Index?.Value is { } lineIndex && lineIndex > 0
                ? reference
                : null;

            ArgbColor color;
            if (outline?.GetFirstChild<Dr.SolidFill>() is { } solidFill)
            {
                if (!colors.TryResolve(solidFill, placeholder: null, out color))
                {
                    return null;
                }
            }
            else if (outline?.GetFirstChild<Dr.GradientFill>() is not null || outline?.GetFirstChild<Dr.PatternFill>() is not null)
            {
                // グラデーション・模様の線は対応していない。スタイルの線で代用しない。
                return null;
            }
            else if (lineRef is null || !colors.TryResolve(lineRef, placeholder: null, out color))
            {
                return null;
            }

            // @w が無ければスタイルの線の書式(テーマの lnStyleLst)の太さ。スタイルも無ければ、
            // 線の色が明示されている以上「見える枠線がある」とみなし、Excelの既定的な細線に近い1ptを補う。
            var widthPt = outline?.Width?.Value is { } width
                ? Units.EmusToPoints(width)
                : lineRef is not null ? colors.GetLineStyleWidthPt(lineRef.Index!.Value) : 1.0;
            return new ShapeOutline(
                color, widthPt, ReadLineEnd(outline?.GetFirstChild<Dr.HeadEnd>()), ReadLineEnd(outline?.GetFirstChild<Dr.TailEnd>()));
        }

        /// <summary>
        /// 接続線の線を読み取る。接続線は <see cref="ConnectorModel.Outline"/> が null のとき Rendering が既定の黒い線を補うため、
        /// 線を明示的に消している場合(<c>a:ln/a:noFill</c>、または図形のスタイルがあり線の色が決まらない場合)は
        /// 透明の線にして、既定の線と区別する(要件10.16)。
        /// </summary>
        private static ShapeOutline? ReadConnectorOutline(
            Xdr.ShapeProperties? shapeProperties, Xdr.ShapeStyle? style, DrawingColorResolver colors)
        {
            var outline = ReadShapeOutline(shapeProperties, style, colors);
            if (outline is not null)
            {
                return outline;
            }

            var line = shapeProperties?.GetFirstChild<Dr.Outline>();
            var explicitNoLine = line?.GetFirstChild<Dr.NoFill>() is not null || style is not null;
            if (explicitNoLine)
            {
                return new ShapeOutline(ArgbColor.Transparent, 0.0);
            }

            // 色の無い a:ln に矢印だけがある場合、既定の黒い線(Rendering が null のとき補う線と同じ)に矢印を付けて返す。
            // null にすると、線は既定値で描かれるのに矢印だけが黙って消えるため(code-reviewer指摘)。
            var headEnd = ReadLineEnd(line?.GetFirstChild<Dr.HeadEnd>());
            var tailEnd = ReadLineEnd(line?.GetFirstChild<Dr.TailEnd>());
            if (headEnd is null && tailEnd is null)
            {
                return null;
            }

            var widthPt = line?.Width?.Value is { } width ? Units.EmusToPoints(width) : 1.0;
            return new ShapeOutline(ArgbColor.Black, widthPt, headEnd, tailEnd);
        }

        /// <summary>線の端の矢印(<c>a:headEnd</c>/<c>a:tailEnd</c>)を読み取る(要件10.18)。<c>none</c>・未指定は null。</summary>
        private static LineEndStyle? ReadLineEnd(Dr.LineEndPropertiesType? lineEnd)
        {
            var type = lineEnd?.Type?.InnerText switch
            {
                "triangle" => LineEndType.Triangle,
                "stealth" => LineEndType.Stealth,
                "arrow" => LineEndType.Arrow,
                "oval" => LineEndType.Oval,
                "diamond" => LineEndType.Diamond,
                _ => (LineEndType?)null,
            };

            if (type is null)
            {
                return null;
            }

            return new LineEndStyle(type.Value, ReadLineEndSize(lineEnd!.Width?.InnerText), ReadLineEndSize(lineEnd.Length?.InnerText));
        }

        private static LineEndSize ReadLineEndSize(string? value) => value switch
        {
            "sm" => LineEndSize.Small,
            "lg" => LineEndSize.Large,
            _ => LineEndSize.Medium,
        };

        /// <summary>
        /// 図形内テキスト(<c>xdr:txBody</c>)を段落・ラン単位で読み取る(要件10.4)。
        /// 合計文字数が<see cref="MaxShapeTextLength"/>を超えた時点で即座に打ち切り、
        /// <paramref name="textTooLong"/>を立てて返す(security-reviewer指摘)。
        /// フォント・色の解析(<see cref="ReadShapeRunFont"/>)は上限を超えていないランに対してのみ
        /// 行うため、極端に大量のランを仕込んだ入力でも処理コストが合計文字数の上限で頭打ちになる。
        /// </summary>
        private static ShapeTextBody? ReadShapeText(
            Xdr.TextBody? textBody, Xdr.ShapeStyle? style, DrawingColorResolver colors, out bool textTooLong)
        {
            textTooLong = false;
            if (textBody is null)
            {
                return null;
            }

            // 文字の a:rPr に色が無ければ、図形のスタイルの fontRef の色を使う(要件10.16)。
            var defaultColor = style?.FontReference is { } fontRef && colors.TryResolve(fontRef, placeholder: null, out var fontColor)
                ? fontColor
                : ArgbColor.Black;

            var paragraphs = new List<ShapeTextParagraph>();
            var totalLength = 0;
            var hasText = false;
            foreach (var paragraph in textBody.Elements<Dr.Paragraph>())
            {
                var runs = new List<ShapeTextRun>();
                foreach (var element in paragraph.ChildElements)
                {
                    // a:r(ラン)と a:fld(フィールド。ページ番号・日付など。保存時点の文字列を a:t に持つ)は
                    // 文字列として、a:br(段落内の改行)は "\n" 1文字のランとして読む。
                    string? text;
                    Dr.RunProperties? runProperties;
                    switch (element)
                    {
                        case Dr.Run run:
                            text = run.Text?.Text;
                            runProperties = run.RunProperties;
                            break;
                        case Dr.Field field:
                            text = field.GetFirstChild<Dr.Text>()?.Text;
                            runProperties = field.GetFirstChild<Dr.RunProperties>();
                            break;
                        case Dr.Break lineBreak:
                            text = "\n";
                            runProperties = lineBreak.GetFirstChild<Dr.RunProperties>();
                            break;
                        default:
                            continue;
                    }

                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    totalLength += text!.Length;
                    if (totalLength > MaxShapeTextLength)
                    {
                        textTooLong = true;
                        return null;
                    }

                    hasText = true;
                    runs.Add(new ShapeTextRun(text!, ReadShapeRunFont(runProperties, defaultColor, colors)));
                }

                // ランの無い段落(空行)も、行を占める空の段落として残す(Runs が空)。
                var hAlign = MapHorizontalAlignment(paragraph.ParagraphProperties?.Alignment?.Value);
                paragraphs.Add(new ShapeTextParagraph(runs, hAlign));
            }

            // 文字が1つも無いテキスト(Excel は文字の無い図形にも空の段落を1つ書く)は、従来どおりテキストなしとする。
            if (!hasText)
            {
                return null;
            }

            var vAlign = MapVerticalAlignment(textBody.BodyProperties?.Anchor?.Value);
            return new ShapeTextBody(paragraphs, vAlign);
        }

        private static FontStyle ReadShapeRunFont(Dr.RunProperties? runProperties, ArgbColor defaultColor, DrawingColorResolver colors)
        {
            if (runProperties is null)
            {
                return FontStyle.Default with { Color = defaultColor };
            }

            // sz は 1/100pt 単位。ST_TextFontSize の範囲(1〜4000pt)外は既定のサイズに戻す(要件6.9)。
            var sizePt = runProperties.FontSize?.Value is { } sz && sz >= MinShapeFontSizeHundredths && sz <= MaxShapeFontSizeHundredths
                ? sz / 100.0
                : FontStyle.Default.SizePt;
            var bold = runProperties.Bold?.Value ?? false;
            var italic = runProperties.Italic?.Value ?? false;
            var underline = MapUnderline(runProperties.Underline?.Value);
            var strike = runProperties.Strike?.Value is { } strikeValue && strikeValue != Dr.TextStrikeValues.NoStrike;
            var name = runProperties.GetFirstChild<Dr.LatinFont>()?.Typeface?.Value ?? FontStyle.Default.Name;
            var color = runProperties.GetFirstChild<Dr.SolidFill>() is { } solidFill
                && colors.TryResolve(solidFill, placeholder: null, out var runColor)
                ? runColor
                : defaultColor;

            return new FontStyle(name!, sizePt, bold, italic, underline, strike, color);
        }

        private static UnderlineStyle MapUnderline(Dr.TextUnderlineValues? value)
        {
            if (value is null || value == Dr.TextUnderlineValues.None)
            {
                return UnderlineStyle.None;
            }

            // 波線・二重下線以外の下線種別は単純な実線として近似する。
            return value == Dr.TextUnderlineValues.Double ? UnderlineStyle.Double : UnderlineStyle.Single;
        }

        private static HorizontalAlignment MapHorizontalAlignment(Dr.TextAlignmentTypeValues? value) => value switch
        {
            Dr.TextAlignmentTypeValues.Center => HorizontalAlignment.Center,
            Dr.TextAlignmentTypeValues.Right => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Left,
        };

        private static VerticalAlignment MapVerticalAlignment(Dr.TextAnchoringTypeValues? value) => value switch
        {
            Dr.TextAnchoringTypeValues.Center => VerticalAlignment.Center,
            Dr.TextAnchoringTypeValues.Bottom => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Top,
        };

        /// <summary>
        /// <paramref name="source"/> から最大 <paramref name="maxBytes"/> バイトだけ読み取る。
        /// 超過した場合は <c>false</c> を返す(要件9で読み取るバイト数に上限を設けるため)。
        /// </summary>
        private static bool TryReadBounded(Stream source, long maxBytes, out byte[] data)
        {
            using var buffer = new MemoryStream();
            var chunk = new byte[81920];
            int read;
            while ((read = source.Read(chunk, 0, chunk.Length)) > 0)
            {
                if (buffer.Length + read > maxBytes)
                {
                    data = Array.Empty<byte>();
                    return false;
                }

                buffer.Write(chunk, 0, read);
            }

            data = buffer.ToArray();
            return true;
        }

        /// <summary>
        /// 画像バイナリの先頭シグネチャ(マジックバイト)が、申告された <paramref name="contentType"/> と
        /// 一致するかを確認する。<see cref="SupportedImageContentTypes"/> に含まれる4形式のみ対応する。
        /// </summary>
        private static bool MatchesContentTypeSignature(string contentType, byte[] data) => contentType.ToLowerInvariant() switch
        {
            "image/png" => data.Length >= 8
                && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
                && data[4] == 0x0D && data[5] == 0x0A && data[6] == 0x1A && data[7] == 0x0A,
            "image/jpeg" => data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF,
            "image/gif" => data.Length >= 6
                && data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38
                && (data[4] == 0x37 || data[4] == 0x39) && data[5] == 0x61,
            "image/bmp" => data.Length >= 2 && data[0] == 0x42 && data[1] == 0x4D,
            _ => false,
        };

        /// <summary>
        /// <c>xdr:from</c>/<c>xdr:to</c> のマーカー(0始まりの行/列 + セル内オフセット)を、
        /// Utsushiの1始まりの <see cref="CellAddress"/> とポイント単位のオフセットに変換する。
        /// </summary>
        private static bool TryReadMarker(Xdr.MarkerType marker, out CellAddress cell, out PointPt offset)
        {
            cell = default;
            offset = default;

            if (marker.ColumnId?.Text is not { } columnText
                || marker.RowId?.Text is not { } rowText
                || !int.TryParse(columnText, NumberStyles.None, CultureInfo.InvariantCulture, out var zeroBasedColumn)
                || !int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out var zeroBasedRow))
            {
                return false;
            }

            var column = zeroBasedColumn + 1;
            var row = zeroBasedRow + 1;
            if (column < 1 || column > CellAddress.MaxColumn || row < 1 || row > CellAddress.MaxRow)
            {
                return false;
            }

            var offsetXEmu = marker.ColumnOffset?.Text is { } colOffsetText
                && long.TryParse(colOffsetText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var x) ? x : 0L;
            var offsetYEmu = marker.RowOffset?.Text is { } rowOffsetText
                && long.TryParse(rowOffsetText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var y) ? y : 0L;

            cell = new CellAddress(row, column);
            offset = new PointPt(Units.EmusToPoints(offsetXEmu), Units.EmusToPoints(offsetYEmu));
            return true;
        }

        private static PageSetupModel ReadPageSetup(
            string sheetName,
            X.Worksheet worksheet,
            IReadOnlyDictionary<string, Dictionary<string, string>> definedNames,
            string? reportCode)
        {
            var setup = worksheet.GetFirstChild<X.PageSetup>();
            var margins = worksheet.GetFirstChild<X.PageMargins>();
            var sheetProperties = worksheet.GetFirstChild<X.SheetProperties>();

            var paper = PaperSizeTable.Resolve(setup?.PaperSize?.Value is { } code ? (int)code : null);
            var orientation = MapOrientation(setup?.Orientation);

            // NaN・無限大・負・極端な値の余白は、その辺だけ Excel の既定値に戻す(要件6.9。NaN の余白で原点が NaN になり
            // 白紙の PDF が出ていた)。
            var pageMargins = margins is null
                ? PageMargins.Default
                : new PageMargins(
                    Units.InchesToPoints(ValidMarginOrDefault(margins.Left?.Value, 0.7)),
                    Units.InchesToPoints(ValidMarginOrDefault(margins.Right?.Value, 0.7)),
                    Units.InchesToPoints(ValidMarginOrDefault(margins.Top?.Value, 0.75)),
                    Units.InchesToPoints(ValidMarginOrDefault(margins.Bottom?.Value, 0.75)),
                    Units.InchesToPoints(ValidMarginOrDefault(margins.Header?.Value, 0.3)),
                    Units.InchesToPoints(ValidMarginOrDefault(margins.Footer?.Value, 0.3)));

            var scaling = ReadScaling(setup, sheetProperties);
            var pageOrder = setup?.PageOrder is not null && setup.PageOrder.Value == X.PageOrderValues.OverThenDown
                ? PageOrder.OverThenDown
                : PageOrder.DownThenOver;

            var rowBreaks = ReadBreaks(worksheet.GetFirstChild<X.RowBreaks>(), CellAddress.MaxRow, sheetName, reportCode);
            var columnBreaks = ReadBreaks(worksheet.GetFirstChild<X.ColumnBreaks>(), CellAddress.MaxColumn, sheetName, reportCode);

            definedNames.TryGetValue(sheetName, out var sheetNames);
            var printAreas = DefinedNameParser.ParsePrintArea(
                sheetNames is not null && sheetNames.TryGetValue(DefinedNameParser.PrintAreaName, out var area) ? area : null,
                MaxPrintAreasPerSheet);
            if (printAreas.Count > MaxPrintAreasPerSheet)
            {
                throw new InvalidExcelFileException(
                    $"シート '{sheetName}' の印刷範囲の個数が上限({MaxPrintAreasPerSheet}個)を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }

            var printTitles = DefinedNameParser.ParsePrintTitles(
                sheetNames is not null && sheetNames.TryGetValue(DefinedNameParser.PrintTitlesName, out var titles) ? titles : null);

            var headerFooter = ReadHeaderFooter(worksheet, sheetName, reportCode);

            return new PageSetupModel(
                paper, orientation, pageMargins, scaling, printAreas, rowBreaks, columnBreaks,
                printTitles, pageOrder, headerFooter);
        }

        /// <summary>ページヘッダー/フッターの設定を読み取る(要件3.7〜3.9)。</summary>
        private static HeaderFooterModel ReadHeaderFooter(X.Worksheet worksheet, string sheetName, string? reportCode)
        {
            var headerFooter = worksheet.GetFirstChild<X.HeaderFooter>();
            if (headerFooter is null)
            {
                return HeaderFooterModel.None;
            }

            foreach (var text in new[]
            {
                headerFooter.OddHeader?.Text, headerFooter.OddFooter?.Text, headerFooter.EvenHeader?.Text,
                headerFooter.EvenFooter?.Text, headerFooter.FirstHeader?.Text, headerFooter.FirstFooter?.Text,
            })
            {
                EnsureTextLengthWithinLimit(
                    text ?? string.Empty, MaxHeaderFooterTextLength, $"シート '{sheetName}' のヘッダー/フッター", reportCode);
            }

            return new HeaderFooterModel(
                headerFooter.OddHeader?.Text,
                headerFooter.OddFooter?.Text,
                headerFooter.EvenHeader?.Text,
                headerFooter.EvenFooter?.Text,
                headerFooter.FirstHeader?.Text,
                headerFooter.FirstFooter?.Text,
                headerFooter.DifferentOddEven?.Value ?? false,
                headerFooter.DifferentFirst?.Value ?? false,
                // scaleWithDoc の既定は true(Excel の既定動作)。
                headerFooter.ScaleWithDoc?.Value ?? true);
        }

        private static PageScaling ReadScaling(X.PageSetup? setup, X.SheetProperties? sheetProperties)
        {
            var scale = (int)Math.Min(setup?.Scale?.Value ?? 100U, 1000U);
            if (scale <= 0)
            {
                scale = 100;
            }

            // Excel の拡大縮小率は 10〜400%。範囲外の値は Excel と同じく丸める(極端な縮小で1ページに
            // 大量のセルを詰め込ませるのを防ぐ。要件6.9。security-reviewer指摘)。
            scale = Math.Max(MinPrintScalePercent, Math.Min(MaxPrintScalePercent, scale));

            // fitToPage が有効なときだけ fitToWidth/fitToHeight を採用する。
            // 値 0 は「その方向は制限しない」を意味するため null に落とす。
            var fitToPage = sheetProperties?.PageSetupProperties?.FitToPage?.Value ?? false;
            if (!fitToPage)
            {
                return new PageScaling(scale, null, null);
            }

            var fitToWidth = (int)(setup?.FitToWidth?.Value ?? 1U);
            var fitToHeight = (int)(setup?.FitToHeight?.Value ?? 1U);

            return new PageScaling(
                scale,
                fitToWidth > 0 ? fitToWidth : null,
                fitToHeight > 0 ? fitToHeight : null);
        }

        private static List<int> ReadBreaks(X.PageBreakType? breaks, int maxIndex, string sheetName, string? reportCode)
        {
            var result = new List<int>();
            if (breaks is null)
            {
                return result;
            }

            var count = 0;
            foreach (var brk in breaks.Elements<X.Break>())
            {
                if (++count > MaxPageBreaksPerSheet)
                {
                    throw new InvalidExcelFileException(
                        $"シート '{sheetName}' の改ページの件数が上限({MaxPageBreaksPerSheet}件)を超えています。",
                        InvalidExcelFileReason.TooLarge,
                        reportCode);
                }

                // man="true" の改ページのみが手動改ページ。自動改ページはレイアウト側で計算する(要件3.2/3.3)。
                if (brk.ManualPageBreak?.Value != true)
                {
                    continue;
                }

                // brk/@id は 0 始まりの行(列)番号で、その行の上(列の左)で改ページする(ECMA-376 Part 1, 18.3.1.3)。
                // つまり id="20" は 1 始まりの 20 行目と 21 行目の間(Excel で 21 行目を選んで改ページを挿入すると id="20")。モデルは「この1始まりの番号の手前で
                // 改ページ」を表すため id + 1 にする。id=0(1行目の上)と、最終行(列)より後ろを指すものは意味が無いため無視する。
                if (brk.Id?.Value is { } id && id >= 1 && id + 1L <= maxIndex)
                {
                    result.Add((int)id + 1);
                }
            }

            result.Sort();
            return result;
        }

        private static PageOrientation MapOrientation(EnumValue<X.OrientationValues>? orientation)
        {
            if (orientation is null)
            {
                return PageOrientation.Portrait;
            }

            if (orientation.Value == X.OrientationValues.Landscape)
            {
                return PageOrientation.Landscape;
            }

            return PageOrientation.Portrait;
        }

        /// <summary>
        /// ブック全体に対するサポート外要素(外部参照)を検出する。要件1.5。
        /// </summary>
        private static void DetectUnsupportedWorkbookElements(WorkbookPart workbookPart, WorkbookReadOptions options)
        {
            if (options.UnsupportedElementBehavior != UnsupportedElementBehavior.Error)
            {
                return;
            }

            if (workbookPart.ExternalWorkbookParts.Any())
            {
                throw new UnsupportedWorkbookElementException(
                    "ブックに外部ブックへの参照(external reference)が含まれています。"
                    + "帳票定義の unsupportedElements が 'error' のため中止します。",
                    "ExternalReference",
                    options.ReportCode);
            }
        }

        /// <summary>
        /// シート単位のサポート外要素(図形・グラフ)を検出する。要件1.5。
        /// </summary>
        private static void DetectUnsupportedElements(
            string sheetName, WorksheetPart worksheetPart, X.Worksheet worksheet, WorkbookReadOptions options)
        {
            if (options.UnsupportedElementBehavior != UnsupportedElementBehavior.Error)
            {
                return;
            }

            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart is not null && drawingsPart.ChartParts.Any())
            {
                throw new UnsupportedWorkbookElementException(
                    $"シート '{sheetName}' にグラフ(chart)が含まれています。帳票定義の unsupportedElements が 'error' のため中止します。",
                    "Chart",
                    options.ReportCode,
                    sheetName);
            }

            // 画像(xdr:pic)は要件9として、シェイプ(xdr:sp)・接続線(xdr:cxnSp)・グループ(xdr:grpSp)は
            // 要件10として常に読み取り対象とするため、ここでは対象外とする(ReadDrawingObjectsが担う)。
            // これら4種の判定は構造的なもの(要素の種類がこの4種かどうか)で、プリセットが
            // 対応済みかどうかは問わない(非対応プリセットはReadShape/ReadConnectorが個別に
            // UnsupportedShapePresetとして検出する。グループ内部の非対応要素はReadGroupShapeが
            // 再帰的に検出しグループ全体を拒否する。要件10.10)。
            // 上記4種のいずれでもない描画オブジェクト(図表枠xdr:graphicFrame等)が1つでもあれば
            // 引き続きサポート外要素とする。drawingsPartが解決できないのに<drawing>参照だけが
            // ある場合は中身を判定できないため、従来どおり保守的にサポート外として扱う。
            var hasUnsupportedDrawingObject = drawingsPart is not null
                ? HasUnsupportedDrawingObject(drawingsPart)
                : worksheet.GetFirstChild<X.Drawing>() is not null;

            if (hasUnsupportedDrawingObject)
            {
                throw new UnsupportedWorkbookElementException(
                    $"シート '{sheetName}' に図形/画像(drawing)が含まれています。帳票定義の unsupportedElements が 'error' のため中止します。",
                    "Drawing",
                    options.ReportCode,
                    sheetName);
            }

            if (worksheet.GetFirstChild<X.LegacyDrawing>() is not null)
            {
                throw new UnsupportedWorkbookElementException(
                    $"シート '{sheetName}' にコメント等のレガシー図形が含まれています。帳票定義の unsupportedElements が 'error' のため中止します。",
                    "LegacyDrawing",
                    options.ReportCode,
                    sheetName);
            }
        }

        /// <summary>
        /// <paramref name="drawingsPart"/> 内のアンカーに、画像(<c>xdr:pic</c>)・シェイプ
        /// (<c>xdr:sp</c>)・接続線(<c>xdr:cxnSp</c>)・グループ(<c>xdr:grpSp</c>)以外の
        /// 描画オブジェクト(図表枠・絶対座標アンカー等)が1つでも含まれるかどうかを判定する
        /// (要件9.5, 10.7)。この4種の判定は構造的なもので、プリセットジオメトリが対応済みか
        /// どうかは問わない(非対応プリセットは<see cref="ReadShape"/>/<see cref="ReadConnector"/>
        /// が個別に検出し、グループ内部の非対応要素は<see cref="ReadGroupShape"/>が再帰的に検出する)。
        /// </summary>
        private static bool HasUnsupportedDrawingObject(DrawingsPart drawingsPart)
        {
            var drawing = drawingsPart.WorksheetDrawing;
            if (drawing is null)
            {
                return false;
            }

            foreach (var anchor in drawing.ChildElements)
            {
                var isSupportedDrawingObjectType = anchor switch
                {
                    Xdr.TwoCellAnchor two => two.GetFirstChild<Xdr.Picture>() is not null
                        || two.GetFirstChild<Xdr.Shape>() is not null
                        || two.GetFirstChild<Xdr.ConnectionShape>() is not null
                        || two.GetFirstChild<Xdr.GroupShape>() is not null,
                    Xdr.OneCellAnchor one => one.GetFirstChild<Xdr.Picture>() is not null
                        || one.GetFirstChild<Xdr.Shape>() is not null
                        || one.GetFirstChild<Xdr.ConnectionShape>() is not null
                        || one.GetFirstChild<Xdr.GroupShape>() is not null,
                    // AbsoluteAnchor(絶対座標配置)は要件9.1/9.2/10.1の対象外のため、対応済み種別でもサポート外として扱う。
                    _ => false,
                };

                if (!isSupportedDrawingObjectType)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<string> ReadSharedStrings(WorkbookPart workbookPart, string? reportCode) =>
            ReadSharedStrings(workbookPart, MaxSharedStringCount, reportCode);

        /// <summary>
        /// <see cref="ReadSharedStrings(WorkbookPart, string?)"/>の本体。実際の上限
        /// (<see cref="MaxSharedStringCount"/>、既定20万件)は現実的なユニットテストでは
        /// 大量の共有文字列を用意しないと到達できないため、<paramref name="maxCount"/>を
        /// 明示的に指定できる形にしてテスト可能にしている。
        /// </summary>
        internal static List<string> ReadSharedStrings(WorkbookPart workbookPart, int maxCount, string? reportCode)
        {
            var result = new List<string>();
            var table = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (table is null)
            {
                return result;
            }

            foreach (var item in table.Elements<X.SharedStringItem>())
            {
                if (result.Count >= maxCount)
                {
                    throw new InvalidExcelFileException(
                        $"共有文字列の数が上限({maxCount}件)を超えています。",
                        InvalidExcelFileReason.TooLarge,
                        reportCode);
                }

                // リッチテキスト(複数 run)の場合は run を連結する。run 単位の書式差は再現しない。
                var text = item.Text?.Text ?? string.Concat(item.Elements<X.Run>().Select(r => r.Text?.Text ?? string.Empty));
                EnsureTextLengthWithinLimit(text, MaxCellTextLength, "共有文字列", reportCode);
                result.Add(text);
            }

            return result;
        }

        /// <summary>シート名 → (定義名 → 参照文字列) の辞書を作る。</summary>
        private static Dictionary<string, Dictionary<string, string>> ReadDefinedNames(WorkbookPart workbookPart)
        {
            var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
            var definedNames = workbookPart.Workbook.DefinedNames;
            if (definedNames is null)
            {
                return result;
            }

            var sheetsByIndex = (workbookPart.Workbook.Sheets?.Elements<X.Sheet>() ?? Enumerable.Empty<X.Sheet>())
                .Select(s => s.Name?.Value ?? string.Empty)
                .ToList();

            foreach (var definedName in definedNames.Elements<X.DefinedName>())
            {
                var name = definedName.Name?.Value;
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                // localSheetId はブック内のシート並び順(0始まり)を指す。
                var localSheetId = definedName.LocalSheetId?.Value;
                var sheetName = localSheetId is { } index && index < sheetsByIndex.Count
                    ? sheetsByIndex[(int)index]
                    : ExtractSheetName(definedName.InnerText);

                if (string.IsNullOrEmpty(sheetName))
                {
                    continue;
                }

                if (!result.TryGetValue(sheetName!, out var perSheet))
                {
                    perSheet = new Dictionary<string, string>(StringComparer.Ordinal);
                    result[sheetName!] = perSheet;
                }

                perSheet[name!] = definedName.InnerText;
            }

            return result;
        }

        /// <summary>"Sheet1!$A$1:$C$3" からシート名部分を取り出す。</summary>
        private static string? ExtractSheetName(string? reference)
        {
            if (string.IsNullOrEmpty(reference))
            {
                return null;
            }

            var index = reference!.IndexOf('!');
            if (index <= 0)
            {
                return null;
            }

            var name = reference.Substring(0, index).Trim();
            if (name.Length >= 2 && name[0] == '\'' && name[name.Length - 1] == '\'')
            {
                name = name.Substring(1, name.Length - 2).Replace("''", "'");
            }

            return name;
        }

        /// <summary>
        /// セルの番地を決める。<c>c/@r</c> は省略できる(ECMA-376 Part 1, 18.3.1.4)。省略時は同じ行の直前のセルの
        /// 次の列(行の先頭なら A 列)とする。<c>r</c> があるのに番地として解釈できないセルは読み飛ばす。
        /// </summary>
        private static bool TryResolveAddress(X.Cell cell, int rowIndex, int previousColumn, out CellAddress address)
        {
            address = default;
            if (cell.CellReference?.Value is { } reference)
            {
                return CellAddress.TryParse(reference, out address);
            }

            var column = previousColumn + 1;
            if (column > CellAddress.MaxColumn || rowIndex < 1 || rowIndex > CellAddress.MaxRow)
            {
                return false;
            }

            address = new CellAddress(rowIndex, column);
            return true;
        }

        /// <summary>行の高さ(pt)が有限かつ 0〜<see cref="MaxRowHeightPt"/> なら返し、そうでなければ null を返す(要件6.9)。</summary>
        internal static double? ValidRowHeightOrNull(double? heightPt) =>
            heightPt is { } h && IsFiniteInRange(h, 0.0, MaxRowHeightPt) ? h : null;

        /// <summary>列幅(文字数)が有限かつ 0〜<see cref="MaxColumnWidthChars"/> なら返し、そうでなければ null を返す(要件6.9)。</summary>
        internal static double? ValidColumnWidthOrNull(double? widthChars) =>
            widthChars is { } w && IsFiniteInRange(w, 0.0, MaxColumnWidthChars) ? w : null;

        /// <summary>
        /// 余白(インチ)が有限かつ 0〜<see cref="MaxPageMarginInches"/> ならその値を、そうでなければ
        /// <paramref name="fallbackInches"/>(Excel の既定値)を返す(要件6.9)。
        /// </summary>
        internal static double ValidMarginOrDefault(double? inches, double fallbackInches) =>
            inches is { } v && IsFiniteInRange(v, 0.0, MaxPageMarginInches) ? v : fallbackInches;

        /// <summary>
        /// フォントサイズ(pt)が有限かつ <see cref="MinFontSizePt"/>〜<see cref="MaxFontSizePt"/> ならその値を、
        /// そうでなければ既定のサイズ(<see cref="FontStyle.Default"/>)を返す(要件6.9)。
        /// </summary>
        internal static double ValidFontSizeOrDefault(double? sizePt) =>
            sizePt is { } v && IsFiniteInRange(v, MinFontSizePt, MaxFontSizePt) ? v : FontStyle.Default.SizePt;

        /// <summary>NaN・無限大でなく、<paramref name="min"/> 以上 <paramref name="max"/> 以下なら true。</summary>
        internal static bool IsFiniteInRange(double value, double min, double max) =>
            !double.IsNaN(value) && !double.IsInfinity(value) && value >= min && value <= max;

        private static void EnsureSize(List<double> list, int oneBasedIndex, double fillValue)
        {
            while (list.Count < oneBasedIndex)
            {
                list.Add(fillValue);
            }
        }

        private static Stream EnsureSeekable(Stream stream, out bool ownsStream, string? reportCode)
        {
            if (stream.CanSeek)
            {
                ownsStream = false;
                return stream;
            }

            var buffer = new MemoryStream();
            CopyWithSizeLimit(stream, buffer, reportCode);
            buffer.Position = 0;
            ownsStream = true;
            return buffer;
        }

        /// <summary>
        /// <see cref="EnsureSeekable"/>専用。シーク不可ストリームを<see cref="MaxXlsxPackageBytes"/>を
        /// 超えない範囲でのみメモリへ複製する。超過した時点で複製済みバッファは破棄し、
        /// 例外化する(要件6, 7。無制限なメモリ確保を防ぐ安全弁)。
        /// </summary>
        private static void CopyWithSizeLimit(Stream source, MemoryStream destination, string? reportCode) =>
            CopyWithSizeLimit(source, destination, MaxXlsxPackageBytes, reportCode);

        /// <summary>
        /// <see cref="CopyWithSizeLimit(Stream, MemoryStream, string?)"/>の本体。実際の上限
        /// (<see cref="MaxXlsxPackageBytes"/>、既定1GiB)は現実的なユニットテストでは到達できない
        /// 大きさのため、<paramref name="maxBytes"/>を明示的に指定できる形にしてテスト可能にしている。
        /// </summary>
        internal static void CopyWithSizeLimit(Stream source, MemoryStream destination, long maxBytes, string? reportCode)
        {
            var buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
                if (total > maxBytes)
                {
                    throw new InvalidExcelFileException(
                        $"入力ファイルのサイズが上限({maxBytes / (1024 * 1024)}MB)を超えています。",
                        InvalidExcelFileReason.TooLarge,
                        reportCode);
                }

                destination.Write(buffer, 0, read);
            }
        }

        /// <summary>
        /// <c>SpreadsheetDocument.Open</c>で実際にパートを展開する前に、ZIPエントリの宣言サイズ
        /// (中央ディレクトリの展開後サイズ)の合計が<see cref="MaxXlsxPackageBytes"/>を超えていないか
        /// 確認する(要件6, 7)。数百バイトのZIPファイルが展開後に極端に大きくなる、
        /// いわゆる「ZIP爆弾」対策(security-reviewer指摘)。ストリームは末尾へ移動して読み戻すため、
        /// 呼び出し前後で<c>Position</c>を復元する。不正なZIP構造(<see cref="InvalidDataException"/>)は
        /// このメソッドでは判定せず、後続の<c>SpreadsheetDocument.Open</c>失敗時の
        /// <see cref="MapOpenFailure"/>による分類に委ねる(検証ロジックの重複を避けるため)。
        /// </summary>
        private static void GuardPackageSize(Stream stream, string? reportCode) =>
            GuardPackageSize(stream, MaxXlsxPackageBytes, reportCode);

        /// <summary>
        /// <see cref="GuardPackageSize(Stream, string?)"/>の本体。実際の上限
        /// (<see cref="MaxXlsxPackageBytes"/>、既定1GiB)は現実的なユニットテストでは到達できない
        /// 大きさのため、<paramref name="maxBytes"/>を明示的に指定できる形にしてテスト可能にしている。
        /// </summary>
        internal static void GuardPackageSize(Stream stream, long maxBytes, string? reportCode) =>
            GuardPackageSize(stream, maxBytes, MaxZipEntries, MaxRelationshipsPerPackage, reportCode);

        /// <summary>
        /// <see cref="GuardPackageSize(Stream, long, string?)"/> の本体。ZIPエントリ数と、全関係パートの関係の数の合計の
        /// 上限も引数に取る(テスト用)。
        /// </summary>
        internal static void GuardPackageSize(
            Stream stream, long maxBytes, int maxEntries, int maxTotalRelationships, string? reportCode)
        {
            if (!stream.CanSeek)
            {
                return;
            }

            var position = stream.Position;
            try
            {
                // ZipArchive は中央ディレクトリの全エントリをメモリ上のオブジェクトにするため、作る前に
                // 終端レコード(EOCD)の申告件数で上限を検査する(要件6.7)。
                if (TryReadDeclaredZipEntryCount(stream) is { } declared && declared > maxEntries)
                {
                    throw TooManyZipEntries(maxEntries, reportCode);
                }

                stream.Position = position;
                ZipArchive archive;
                IReadOnlyCollection<ZipArchiveEntry> entries;
                try
                {
                    archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
                    entries = archive.Entries;
                }
                catch (Exception ex) when (ex is InvalidDataException or IOException)
                {
                    // ZIPでない(暗号化ブック・旧形式など)・中央ディレクトリが壊れている場合は、
                    // SpreadsheetDocument.Open 側の失敗分類(PasswordProtected/NotOpenXmlFormat/Corrupted)に委ねる。
                    return;
                }

                using (archive)
                {
                    // 申告件数と実際の件数が食い違うファイルもあるため、実際の件数でも確かめる。
                    if (entries.Count > maxEntries)
                    {
                        throw TooManyZipEntries(maxEntries, reportCode);
                    }

                    long total = 0;
                    foreach (var entry in entries)
                    {
                        total += entry.Length;
                        if (total > maxBytes)
                        {
                            throw new InvalidExcelFileException(
                                $"入力ファイルの展開後サイズが上限({maxBytes / (1024 * 1024)}MB)を超えています。",
                                InvalidExcelFileReason.TooLarge,
                                reportCode);
                        }
                    }

                    // [Content_Types].xml はパートではないため GuardXmlParts の対象にならず、Open の中で DOM として
                    // 読まれる。Open より前に大きさ・要素数・深さを検査する(要件6.7)。
                    foreach (var entry in entries)
                    {
                        if (string.Equals(entry.FullName, ContentTypesEntryName, StringComparison.OrdinalIgnoreCase))
                        {
                            GuardZipEntry(entry, reportCode, s => EnsureContentTypesWithinLimits(
                                s, MaxContentTypesBytes, MaxContentTypesEntries, entry.FullName, reportCode));
                        }
                    }

                    // 関係パートは SpreadsheetDocument.Open の中で解析されるため、Open より前に件数を検査する(要件6.7)。
                    // 展開できない関係パートが1つあっても後ろの関係パートの検査を飛ばさないよう、エントリごとに扱う。
                    var totalRelationships = 0;
                    foreach (var entry in entries)
                    {
                        if (!entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        GuardZipEntry(entry, reportCode, s =>
                        {
                            totalRelationships += EnsureRelationshipCountWithinLimit(
                                s, MaxRelationshipsPerPart, entry.FullName, reportCode);
                        });

                        if (totalRelationships > maxTotalRelationships)
                        {
                            throw new InvalidExcelFileException(
                                $"関係パート全体の関係の数が上限({maxTotalRelationships}件)を超えています。",
                                InvalidExcelFileReason.TooLarge,
                                reportCode);
                        }
                    }
                }
            }
            finally
            {
                stream.Position = position;
            }
        }

        /// <summary>
        /// ZIPエントリを開いて <paramref name="inspect"/> で検査する。圧縮データが壊れていて展開できないエントリは
        /// Corrupted にする(以前は検査全体を打ち切っていたため、後ろのエントリが検査されなかった。security-reviewer指摘)。
        /// </summary>
        private static void GuardZipEntry(ZipArchiveEntry entry, string? reportCode, Action<Stream> inspect)
        {
            try
            {
                using var entryStream = entry.Open();
                inspect(entryStream);
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException)
            {
                throw new InvalidExcelFileException(
                    $"エントリ '{entry.FullName}' の圧縮データが壊れているため読み取れません。",
                    InvalidExcelFileReason.Corrupted,
                    reportCode,
                    ex);
            }
        }

        private static InvalidExcelFileException TooManyZipEntries(int maxEntries, string? reportCode) =>
            new(
                $"入力ファイル(ZIP)のエントリの数が上限({maxEntries}個)を超えています。",
                InvalidExcelFileReason.TooLarge,
                reportCode);

        /// <summary>
        /// ZIPの終端レコード(EOCD。ZIP64 なら ZIP64 EOCD)が申告するエントリ数を読む。見つからない・読めない場合は null
        /// (その判断は ZipArchive/Open に委ねる)。ストリームの位置は変わる。
        /// </summary>
        internal static long? TryReadDeclaredZipEntryCount(Stream stream)
        {
            const int EocdMinLength = 22;
            const int MaxCommentLength = 0xFFFF;
            var length = stream.Length;
            if (length < EocdMinLength)
            {
                return null;
            }

            var tailLength = (int)Math.Min(length, EocdMinLength + MaxCommentLength);
            var tail = new byte[tailLength];
            stream.Position = length - tailLength;
            var read = 0;
            while (read < tailLength)
            {
                var n = stream.Read(tail, read, tailLength - read);
                if (n <= 0)
                {
                    return null;
                }

                read += n;
            }

            for (var i = tailLength - EocdMinLength; i >= 0; i--)
            {
                if (tail[i] != 0x50 || tail[i + 1] != 0x4B || tail[i + 2] != 0x05 || tail[i + 3] != 0x06)
                {
                    continue;
                }

                long count = tail[i + 10] | (tail[i + 11] << 8);
                if (count != 0xFFFF)
                {
                    return count;
                }

                // ZIP64: EOCD の直前20バイトに ZIP64 EOCD ロケータ(PK\x06\x07)があり、ZIP64 EOCD の位置を指す。
                var locator = i - 20;
                if (locator < 0 || tail[locator] != 0x50 || tail[locator + 1] != 0x4B
                    || tail[locator + 2] != 0x06 || tail[locator + 3] != 0x07)
                {
                    return count;
                }

                var zip64Offset = BitConverter.ToInt64(tail, locator + 8);
                if (zip64Offset < 0 || zip64Offset + 56 > length)
                {
                    return null;
                }

                var record = new byte[56];
                stream.Position = zip64Offset;
                if (stream.Read(record, 0, record.Length) != record.Length
                    || record[0] != 0x50 || record[1] != 0x4B || record[2] != 0x06 || record[3] != 0x06)
                {
                    return null;
                }

                // ZIP64 EOCD の「全エントリ数」は32バイト目からの8バイト。
                var count64 = BitConverter.ToInt64(record, 32);
                return count64 < 0 ? long.MaxValue : count64;
            }

            return null;
        }

        /// <summary>
        /// <c>[Content_Types].xml</c> の大きさ・要素の数・入れ子の深さが上限以下か、流し読みで確認する(要件6.7。
        /// テスト用に上限を引数に取る)。XMLとして壊れている場合は <c>SpreadsheetDocument.Open</c> 側の失敗分類に委ねる。
        /// </summary>
        internal static void EnsureContentTypesWithinLimits(
            Stream contentTypes, long maxBytes, int maxEntries, string partName, string? reportCode)
        {
            try
            {
                using var bounded = new BoundedReadStream(contentTypes, maxBytes);
                using var reader = XmlReader.Create(bounded, CreateGuardReaderSettings());
                var count = 0;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element || reader.Depth == 0)
                    {
                        continue;
                    }

                    // ルート(Types)の子の Default/Override を数える。それより深い要素は本来無いが、深い入れ子で
                    // DOM の組み立てに時間・スタックを使わせないよう、同じく数え、深さも制限する。
                    if (++count > maxEntries)
                    {
                        throw new InvalidExcelFileException(
                            $"'{partName}' の要素(Default/Override)の数が上限({maxEntries}個)を超えています。",
                            InvalidExcelFileReason.TooLarge,
                            reportCode);
                    }

                    if (reader.Depth >= MaxXmlElementDepth)
                    {
                        throw new InvalidExcelFileException(
                            $"'{partName}' の要素の入れ子が上限({MaxXmlElementDepth}段)を超えています。",
                            InvalidExcelFileReason.TooLarge,
                            reportCode);
                    }
                }
            }
            catch (BoundedReadStream.LimitExceededException)
            {
                throw new InvalidExcelFileException(
                    $"'{partName}' の大きさが上限({maxBytes / 1024}KB)を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }
            catch (XmlException)
            {
                // 壊れた [Content_Types].xml は Open 側で Corrupted として分類される。
            }
        }

        private static XmlReaderSettings CreateGuardReaderSettings() => new()
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            IgnoreProcessingInstructions = true,
        };

        /// <summary>
        /// 関係パート(<c>*.rels</c>)の関係の数が上限以下か、流し読みで確認し、数えた関係の数を返す(要件6.7。
        /// テスト用に上限を引数に取る)。XMLとして壊れている場合は、ここでは判断せず <c>SpreadsheetDocument.Open</c> 側の
        /// 失敗分類に委ねる(それまでに数えた数を返す)。
        /// </summary>
        internal static int EnsureRelationshipCountWithinLimit(Stream rels, int maxRelationships, string partName, string? reportCode)
        {
            var count = 0;
            try
            {
                using var bounded = new BoundedReadStream(rels, MaxXmlPartBytes);
                using var reader = XmlReader.Create(bounded, CreateGuardReaderSettings());
                while (reader.Read())
                {
                    // ルート要素(Relationships)は数えず、子の Relationship だけを数える。
                    if (reader.NodeType == XmlNodeType.Element && reader.Depth > 0 && ++count > maxRelationships)
                    {
                        throw new InvalidExcelFileException(
                            $"関係パート '{partName}' の関係の数が上限({maxRelationships}件)を超えています。",
                            InvalidExcelFileReason.TooLarge,
                            reportCode);
                    }
                }
            }
            catch (BoundedReadStream.LimitExceededException)
            {
                throw new InvalidExcelFileException(
                    $"関係パート '{partName}' の大きさが上限({MaxXmlPartBytes / (1024 * 1024)}MB)を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }
            catch (XmlException)
            {
                // 壊れた関係パートは Open 側で Corrupted として分類される。
            }

            return count;
        }

        /// <summary>
        /// パッケージ内の全XMLパートを <see cref="XmlReader"/> で流し読みし、要素の入れ子の深さと
        /// 展開後の大きさが上限以下か確認する(要件6.7)。OpenXml SDK が DOM を組み立てる前に呼ぶこと。
        /// 流し読みは再帰しないため、深いネストでもスタックを消費しない。
        /// </summary>
        private static void GuardXmlParts(SpreadsheetDocument document, string? reportCode) =>
            GuardXmlParts(document, MaxXmlElementsPerPackage, reportCode);

        /// <summary>
        /// <see cref="GuardXmlParts(SpreadsheetDocument, string?)"/> の本体。パッケージ全体の要素の数の上限を引数に取る(テスト用)。
        /// パートごとの上限だけでは、上限近くのパートを複数並べて DOM のメモリを積み上げられるため、全パートの合計にも
        /// 上限を設ける(security-reviewer指摘)。
        /// </summary>
        internal static void GuardXmlParts(SpreadsheetDocument document, long maxElementsPerPackage, string? reportCode)
        {
            var total = 0L;
            foreach (var part in document.GetAllParts())
            {
                if (!IsXmlContentType(part.ContentType))
                {
                    continue;
                }

                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                total += EnsureXmlWithinLimits(
                    stream,
                    MaxXmlElementDepth,
                    MaxXmlPartBytes,
                    MaxXmlElementsPerPart,
                    maxElementsPerPackage - total,
                    maxElementsPerPackage,
                    part.Uri.ToString(),
                    reportCode);
            }
        }

        private static bool IsXmlContentType(string contentType) =>
            contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase)
            || contentType.EndsWith("/xml", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// XMLストリームの要素の入れ子の深さと大きさが上限以下か確認する(要件6.7。テスト用に上限を引数に取る)。
        /// </summary>
        internal static long EnsureXmlWithinLimits(Stream xml, int maxDepth, long maxBytes, string partName, string? reportCode) =>
            EnsureXmlWithinLimits(xml, maxDepth, maxBytes, MaxXmlElementsPerPart, partName, reportCode);

        /// <summary>
        /// <see cref="EnsureXmlWithinLimits(Stream, int, long, string, string?)"/> の本体。要素の数の上限も引数に取る(テスト用)。
        /// 数えた要素の数を返す。
        /// </summary>
        internal static long EnsureXmlWithinLimits(
            Stream xml, int maxDepth, long maxBytes, long maxElements, string partName, string? reportCode) =>
            EnsureXmlWithinLimits(xml, maxDepth, maxBytes, maxElements, long.MaxValue, long.MaxValue, partName, reportCode);

        private static long EnsureXmlWithinLimits(
            Stream xml,
            int maxDepth,
            long maxBytes,
            long maxElements,
            long remainingPackageElements,
            long maxPackageElements,
            string partName,
            string? reportCode)
        {
            var elements = 0L;
            try
            {
                using var bounded = new BoundedReadStream(xml, maxBytes);
                using var reader = XmlReader.Create(bounded, CreateGuardReaderSettings());
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }

                    if (++elements > maxElements)
                    {
                        throw new InvalidExcelFileException(
                            $"パート '{partName}' の要素の数が上限({maxElements}個)を超えています。",
                            InvalidExcelFileReason.TooLarge,
                            reportCode);
                    }

                    if (elements > remainingPackageElements)
                    {
                        throw new InvalidExcelFileException(
                            $"入力ファイル全体のXMLの要素の数が上限({maxPackageElements}個)を超えています(パート '{partName}' で超過)。",
                            InvalidExcelFileReason.TooLarge,
                            reportCode);
                    }

                    if (reader.Depth >= maxDepth)
                    {
                        throw new InvalidExcelFileException(
                            $"パート '{partName}' の要素の入れ子が上限({maxDepth}段)を超えています。",
                            InvalidExcelFileReason.TooLarge,
                            reportCode);
                    }
                }
            }
            catch (BoundedReadStream.LimitExceededException)
            {
                throw new InvalidExcelFileException(
                    $"パート '{partName}' の大きさが上限({maxBytes / (1024 * 1024)}MB)を超えています。",
                    InvalidExcelFileReason.TooLarge,
                    reportCode);
            }
            catch (Exception ex) when (ex is XmlException or InvalidDataException or IOException)
            {
                // 壊れた圧縮データ(deflate)は InvalidDataException/IOException になる(code-reviewer指摘)。
                throw new InvalidExcelFileException(
                    $"パート '{partName}' のXMLが壊れているため読み取れません。",
                    InvalidExcelFileReason.Corrupted,
                    reportCode,
                    ex);
            }

            return elements;
        }

        /// <summary>読み取ったバイト数が上限を超えたら例外を投げる読み取り専用ストリーム。</summary>
        private sealed class BoundedReadStream : Stream
        {
            private readonly Stream _inner;
            private readonly long _maxBytes;
            private long _read;

            public BoundedReadStream(Stream inner, long maxBytes)
            {
                _inner = inner;
                _maxBytes = maxBytes;
            }

            public override bool CanRead => true;

            public override bool CanSeek => false;

            public override bool CanWrite => false;

            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => _read;
                set => throw new NotSupportedException();
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                var n = _inner.Read(buffer, offset, count);
                _read += n;
                if (_read > _maxBytes)
                {
                    throw new LimitExceededException();
                }

                return n;
            }

            public override void Flush()
            {
            }

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            /// <summary>上限超過を <see cref="XmlException"/> と区別するための例外。</summary>
            public sealed class LimitExceededException : Exception
            {
            }
        }

        /// <summary>
        /// <c>SpreadsheetDocument.Open</c> の失敗理由を <see cref="InvalidExcelFileException"/> に分類する(要件6.1, 6.2)。
        /// </summary>
        private static InvalidExcelFileException MapOpenFailure(Exception ex, Stream stream, string? reportCode)
        {
            var reason = ClassifyOpenFailure(ex, stream);
            var message = reason switch
            {
                InvalidExcelFileReason.PasswordProtected =>
                    "入力ファイルはパスワード保護(暗号化)されています。Utsushi は暗号化されたブックを読み取れません。",
                InvalidExcelFileReason.NotOpenXmlFormat =>
                    "入力ファイルは .xlsx(Open XML)形式ではありません。.xls / .xlsb などの旧形式は対象外です。",
                InvalidExcelFileReason.Corrupted =>
                    "入力ファイルの Open XML 構造が壊れているため読み取れません。",
                _ => "入力ファイルを Excel ブックとして読み取れませんでした。",
            };

            return new InvalidExcelFileException(message, reason, reportCode, ex);
        }

        private static InvalidExcelFileReason ClassifyOpenFailure(Exception ex, Stream stream)
        {
            // 暗号化された OOXML は OLE 複合ドキュメント(CFB)としてラップされており、
            // 先頭が ZIP シグネチャ "PK" ではなく CFB シグネチャ D0 CF 11 E0 になる。
            var signature = TryReadSignature(stream);
            if (signature is { } sig)
            {
                if (sig.Length >= 4 && sig[0] == 0xD0 && sig[1] == 0xCF && sig[2] == 0x11 && sig[3] == 0xE0)
                {
                    // .xls(BIFF8)も同じシグネチャを持つ。暗号化xlsxか旧形式かは判別できないため
                    // より運用上起こりやすいパスワード保護として報告する。
                    return InvalidExcelFileReason.PasswordProtected;
                }

                if (sig.Length >= 2 && (sig[0] != 0x50 || sig[1] != 0x4B))
                {
                    return InvalidExcelFileReason.NotOpenXmlFormat;
                }
            }

            // 壊れた関係パート(.rels)は Open の中で XmlException になる。
            if (ex is InvalidDataException or FileFormatException or XmlException)
            {
                return InvalidExcelFileReason.Corrupted;
            }

            return ex is OpenXmlPackageException ? InvalidExcelFileReason.Corrupted : InvalidExcelFileReason.Unknown;
        }

        private static byte[]? TryReadSignature(Stream stream)
        {
            if (!stream.CanSeek)
            {
                return null;
            }

            var original = stream.Position;
            try
            {
                stream.Position = 0;
                var buffer = new byte[8];
                var read = stream.Read(buffer, 0, buffer.Length);
                if (read <= 0)
                {
                    return null;
                }

                Array.Resize(ref buffer, read);
                return buffer;
            }
            catch (IOException)
            {
                return null;
            }
            finally
            {
                try
                {
                    stream.Position = original;
                }
                catch (IOException)
                {
                    // 位置を戻せない場合は呼び出し元で例外化済みのためここでは無視する。
                }
            }
        }
    }
}
