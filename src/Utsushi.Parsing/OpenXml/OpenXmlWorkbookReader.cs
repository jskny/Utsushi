using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
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
            var seekable = EnsureSeekable(xlsxStream, out var ownsStream);
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
            SpreadsheetDocument document;
            try
            {
                document = SpreadsheetDocument.Open(stream, isEditable: false);
            }
            catch (Exception ex)
            {
                throw MapOpenFailure(ex, stream, options.ReportCode);
            }

            using (document)
            {
                var workbookPart = document.WorkbookPart
                    ?? throw new InvalidExcelFileException(
                        "ワークブックパートが存在しません。ファイルが破損している可能性があります。",
                        InvalidExcelFileReason.Corrupted,
                        options.ReportCode);

                DetectUnsupportedWorkbookElements(workbookPart, options);

                var colors = ColorResolver.Create(workbookPart);
                var styles = StyleTable.Create(workbookPart, colors);
                var sharedStrings = ReadSharedStrings(workbookPart);
                var definedNames = ReadDefinedNames(workbookPart);

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

                    if (workbookPart.GetPartById(sheet.Id!.Value!) is not WorksheetPart worksheetPart)
                    {
                        continue;
                    }

                    sheets.Add(ReadSheet(name!, worksheetPart, styles, sharedStrings, definedNames, options));
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
            WorkbookReadOptions options)
        {
            var worksheet = worksheetPart.Worksheet;

            DetectUnsupportedElements(name, worksheetPart, worksheet, options);

            var cells = new Dictionary<CellAddress, CellModel>();
            var rowHeights = new List<double>();
            var hiddenRows = new HashSet<int>();

            var sheetFormat = worksheet.GetFirstChild<X.SheetFormatProperties>();
            var defaultRowHeight = sheetFormat?.DefaultRowHeight?.Value ?? 15.0;
            var defaultColumnWidth = ResolveDefaultColumnWidth(sheetFormat);

            var sheetData = worksheet.GetFirstChild<X.SheetData>();
            if (sheetData is not null)
            {
                foreach (var row in sheetData.Elements<X.Row>())
                {
                    var rowIndex = (int)(row.RowIndex?.Value ?? 0U);
                    if (rowIndex < 1)
                    {
                        continue;
                    }

                    EnsureSize(rowHeights, rowIndex, defaultRowHeight);
                    if (row.CustomHeight?.Value == true && row.Height?.Value is { } height)
                    {
                        rowHeights[rowIndex - 1] = height;
                    }

                    if (row.Hidden?.Value == true)
                    {
                        hiddenRows.Add(rowIndex);
                    }

                    foreach (var cell in row.Elements<X.Cell>())
                    {
                        if (!TryResolveAddress(cell, rowIndex, out var address))
                        {
                            continue;
                        }

                        cells[address] = ReadCell(cell, styles, sharedStrings);
                    }
                }
            }

            var (columnWidths, hiddenColumns) = ReadColumns(worksheet, defaultColumnWidth);
            var mergedRanges = ReadMergedRanges(worksheet);
            var pageSetup = ReadPageSetup(name, worksheet, definedNames);
            var images = ReadImages(name, worksheetPart, options);

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
                images.Cast<DrawingObjectModel>().ToList());
        }

        /// <summary>
        /// <c>defaultColWidth</c> が無いブックでは <c>baseColWidth</c>(既定8文字)から既定列幅を導く。
        /// </summary>
        private static double ResolveDefaultColumnWidth(X.SheetFormatProperties? sheetFormat)
        {
            if (sheetFormat?.DefaultColumnWidth?.Value is { } explicitWidth)
            {
                return explicitWidth;
            }

            var baseWidth = sheetFormat?.BaseColumnWidth?.Value ?? 8U;

            // Excel の既定列幅(8.43文字)は baseColWidth=8 にパディングを加えた値に相当する。
            return baseWidth + 0.43;
        }

        private static CellModel ReadCell(X.Cell cell, StyleTable styles, IReadOnlyList<string> sharedStrings)
        {
            var style = styles.GetCellStyle(cell.StyleIndex?.Value is { } s ? (int)s : null);
            var hasFormula = cell.CellFormula is not null;
            var rawValue = cell.CellValue?.InnerText;
            var dataType = cell.DataType?.Value;

            // インライン文字列
            if (dataType is not null && dataType == X.CellValues.InlineString)
            {
                var text = cell.InlineString?.Text?.Text ?? cell.InlineString?.InnerText ?? string.Empty;
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

        private static (List<double> Widths, HashSet<int> Hidden) ReadColumns(X.Worksheet worksheet, double defaultWidth)
        {
            var widths = new List<double>();
            var hidden = new HashSet<int>();

            var columns = worksheet.GetFirstChild<X.Columns>();
            if (columns is null)
            {
                return (widths, hidden);
            }

            foreach (var column in columns.Elements<X.Column>())
            {
                var min = (int)(column.Min?.Value ?? 0U);
                var max = (int)(column.Max?.Value ?? 0U);
                if (min < 1 || max < min)
                {
                    continue;
                }

                // Excel は未使用の右端まで Column 要素を伸ばすことがある。使用範囲を超える定義は既定幅と同じなので無視する。
                max = Math.Min(max, CellAddress.MaxColumn);

                var isHidden = column.Hidden?.Value == true;
                var width = column.Width?.Value ?? defaultWidth;
                var hasCustomWidth = column.CustomWidth?.Value == true || column.Width?.Value is not null;

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

        private static List<MergedRange> ReadMergedRanges(X.Worksheet worksheet)
        {
            var result = new List<MergedRange>();
            var mergeCells = worksheet.GetFirstChild<X.MergeCells>();
            if (mergeCells is null)
            {
                return result;
            }

            foreach (var merge in mergeCells.Elements<X.MergeCell>())
            {
                if (merge.Reference?.Value is { } reference && CellRange.TryParse(reference, out var range))
                {
                    result.Add(new MergedRange(range));
                }
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
        /// シートに埋め込まれた画像(<c>xdr:pic</c>)を読み取る(要件9)。
        /// </summary>
        /// <remarks>
        /// 画像以外の描画オブジェクト(シェイプ・グラフ等)は対象外とし、<see cref="DetectUnsupportedElements"/>
        /// 側で引き続き「サポート外要素」として扱う(要件9.5)。画像は帳票定義の <c>unsupportedElements</c>
        /// 設定によらず常に読み取り対象とするが、デコード不能な形式(EMF/WMF等)・対応形式を偽装した
        /// バイナリ・上限を超える枚数/サイズだけは同じ設定に従う(要件9.4)。
        /// </remarks>
        private static List<ImageModel> ReadImages(string sheetName, WorksheetPart worksheetPart, WorkbookReadOptions options)
        {
            var result = new List<ImageModel>();
            var drawing = worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawing is null)
            {
                return result;
            }

            var drawingsPart = worksheetPart.DrawingsPart!;

            foreach (var anchor in drawing.ChildElements)
            {
                var (fromMarker, picture) = anchor switch
                {
                    Xdr.TwoCellAnchor two => (two.FromMarker, two.GetFirstChild<Xdr.Picture>()),
                    Xdr.OneCellAnchor one => (one.FromMarker, one.GetFirstChild<Xdr.Picture>()),
                    _ => (null, null),
                };

                if (fromMarker is null || picture is null)
                {
                    continue;
                }

                if (!TryReadMarker(fromMarker, out var anchorCell, out var anchorOffset))
                {
                    continue;
                }

                if (result.Count >= MaxImagesPerSheet)
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

                    // ignore時は上限を超えた以降の画像アンカーをまとめて無視する。
                    break;
                }

                var embedId = picture.BlipFill?.Blip?.Embed?.Value;
                if (string.IsNullOrEmpty(embedId) || drawingsPart.GetPartById(embedId!) is not ImagePart imagePart)
                {
                    continue;
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

                    continue;
                }

                AnchorExtent? extent = anchor switch
                {
                    Xdr.TwoCellAnchor two when two.ToMarker is { } toMarker && TryReadMarker(toMarker, out var toCell, out var toOffset) =>
                        new CellSpanAnchorExtent(toCell, toOffset),
                    Xdr.OneCellAnchor { Extent: { Cx: { } cx, Cy: { } cy } } =>
                        new FixedAnchorExtent(Units.EmusToPoints(cx.Value), Units.EmusToPoints(cy.Value)),
                    _ => null,
                };

                if (extent is null)
                {
                    continue;
                }

                byte[] data;
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

                        continue;
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

                    continue;
                }

                result.Add(new ImageModel(data, imagePart.ContentType, anchorCell, anchorOffset, extent));
            }

            return result;
        }

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
            IReadOnlyDictionary<string, Dictionary<string, string>> definedNames)
        {
            var setup = worksheet.GetFirstChild<X.PageSetup>();
            var margins = worksheet.GetFirstChild<X.PageMargins>();
            var sheetProperties = worksheet.GetFirstChild<X.SheetProperties>();

            var paper = PaperSizeTable.Resolve(setup?.PaperSize?.Value is { } code ? (int)code : null);
            var orientation = MapOrientation(setup?.Orientation);

            var pageMargins = margins is null
                ? PageMargins.Default
                : new PageMargins(
                    Units.InchesToPoints(margins.Left?.Value ?? 0.7),
                    Units.InchesToPoints(margins.Right?.Value ?? 0.7),
                    Units.InchesToPoints(margins.Top?.Value ?? 0.75),
                    Units.InchesToPoints(margins.Bottom?.Value ?? 0.75),
                    Units.InchesToPoints(margins.Header?.Value ?? 0.3),
                    Units.InchesToPoints(margins.Footer?.Value ?? 0.3));

            var scaling = ReadScaling(setup, sheetProperties);
            var pageOrder = setup?.PageOrder is not null && setup.PageOrder.Value == X.PageOrderValues.OverThenDown
                ? PageOrder.OverThenDown
                : PageOrder.DownThenOver;

            var rowBreaks = ReadBreaks(worksheet.GetFirstChild<X.RowBreaks>());
            var columnBreaks = ReadBreaks(worksheet.GetFirstChild<X.ColumnBreaks>());

            definedNames.TryGetValue(sheetName, out var sheetNames);
            var printAreas = DefinedNameParser.ParsePrintArea(
                sheetNames is not null && sheetNames.TryGetValue(DefinedNameParser.PrintAreaName, out var area) ? area : null);
            var printTitles = DefinedNameParser.ParsePrintTitles(
                sheetNames is not null && sheetNames.TryGetValue(DefinedNameParser.PrintTitlesName, out var titles) ? titles : null);

            var headerFooter = ReadHeaderFooter(worksheet);

            return new PageSetupModel(
                paper, orientation, pageMargins, scaling, printAreas, rowBreaks, columnBreaks,
                printTitles, pageOrder, headerFooter);
        }

        /// <summary>ページヘッダー/フッターの設定を読み取る(要件3.7〜3.9)。</summary>
        private static HeaderFooterModel ReadHeaderFooter(X.Worksheet worksheet)
        {
            var headerFooter = worksheet.GetFirstChild<X.HeaderFooter>();
            if (headerFooter is null)
            {
                return HeaderFooterModel.None;
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
            var scale = (int)(setup?.Scale?.Value ?? 100U);
            if (scale <= 0)
            {
                scale = 100;
            }

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

        private static List<int> ReadBreaks(X.PageBreakType? breaks)
        {
            var result = new List<int>();
            if (breaks is null)
            {
                return result;
            }

            foreach (var brk in breaks.Elements<X.Break>())
            {
                // man="true" の改ページのみが手動改ページ。自動改ページはレイアウト側で計算する(要件3.2/3.3)。
                if (brk.ManualPageBreak?.Value != true)
                {
                    continue;
                }

                if (brk.Id?.Value is { } id && id > 0)
                {
                    result.Add((int)id);
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

            // 画像(xdr:pic)は要件9として常に読み取り対象とするため、ここでは対象外とする(ReadImagesが担う)。
            // 画像以外の描画オブジェクト(シェイプ・グループ・接続線等)が1つでもあれば引き続きサポート外要素とする。
            // drawingsPartが解決できないのに<drawing>参照だけがある場合は中身を判定できないため、
            // 従来どおり保守的にサポート外として扱う。
            var hasUnsupportedDrawingObject = drawingsPart is not null
                ? HasNonPictureDrawingObject(drawingsPart)
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
        /// <paramref name="drawingsPart"/> 内のアンカーに、画像(<c>xdr:pic</c>)以外の描画オブジェクト
        /// (シェイプ・グループ・接続線・絶対座標アンカー等)が1つでも含まれるかどうかを判定する(要件9.5)。
        /// </summary>
        private static bool HasNonPictureDrawingObject(DrawingsPart drawingsPart)
        {
            var drawing = drawingsPart.WorksheetDrawing;
            if (drawing is null)
            {
                return false;
            }

            foreach (var anchor in drawing.ChildElements)
            {
                var picture = anchor switch
                {
                    Xdr.TwoCellAnchor two => two.GetFirstChild<Xdr.Picture>(),
                    Xdr.OneCellAnchor one => one.GetFirstChild<Xdr.Picture>(),
                    // AbsoluteAnchor(絶対座標配置)は要件9.1/9.2の対象外のため、画像でもサポート外として扱う。
                    _ => null,
                };

                if (picture is null)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<string> ReadSharedStrings(WorkbookPart workbookPart)
        {
            var result = new List<string>();
            var table = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (table is null)
            {
                return result;
            }

            foreach (var item in table.Elements<X.SharedStringItem>())
            {
                // リッチテキスト(複数 run)の場合は run を連結する。run 単位の書式差は再現しない。
                if (item.Text?.Text is { } text)
                {
                    result.Add(text);
                    continue;
                }

                var runs = item.Elements<X.Run>().Select(r => r.Text?.Text ?? string.Empty);
                result.Add(string.Concat(runs));
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

        private static bool TryResolveAddress(X.Cell cell, int rowIndex, out CellAddress address)
        {
            if (cell.CellReference?.Value is { } reference && CellAddress.TryParse(reference, out address))
            {
                return true;
            }

            address = default;
            return false;
        }

        private static void EnsureSize(List<double> list, int oneBasedIndex, double fillValue)
        {
            while (list.Count < oneBasedIndex)
            {
                list.Add(fillValue);
            }
        }

        private static Stream EnsureSeekable(Stream stream, out bool ownsStream)
        {
            if (stream.CanSeek)
            {
                ownsStream = false;
                return stream;
            }

            var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            buffer.Position = 0;
            ownsStream = true;
            return buffer;
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

            if (ex is InvalidDataException or FileFormatException)
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
