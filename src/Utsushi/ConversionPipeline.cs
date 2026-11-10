using System;
using System.Collections.Generic;
using System.IO;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout;
using Utsushi.Layout.Model;
using Utsushi.Parsing;
using Utsushi.Parsing.Model;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi
{
    /// <summary>
    /// 5つのレイヤー(Parsing → 変換設定 → Substitution → Layout → Rendering)をつなぐ変換の手順(excel2pdf-consolidation 要件1.4)。
    /// <see cref="Excel2PdfEngine"/> が持ち、<see cref="Excel2Pdf"/> から使う。
    /// </summary>
    /// <remarks>
    /// 状態を持たず、変換ごとの状態は呼び出しごとに作る。共有するのはフォント・文字幅のキャッシュと、数値書式の解析結果のキャッシュだけである。
    /// </remarks>
    internal sealed class ConversionPipeline : IDisposable
    {
        private readonly IWorkbookReader _workbookReader;
        private readonly IReportModelBuilder _modelBuilder;
        private readonly Substitution.ICellSubstitutor _substitutor;
        private readonly IReportLayoutEngine _layoutEngine;
        private readonly IPdfRenderer _renderer;
        private readonly IDisposable? _ownedResources;

        public ConversionPipeline(
            IWorkbookReader workbookReader,
            IReportModelBuilder modelBuilder,
            Substitution.ICellSubstitutor substitutor,
            IReportLayoutEngine layoutEngine,
            IPdfRenderer renderer,
            IDisposable? ownedResources = null)
        {
            _workbookReader = workbookReader ?? throw new ArgumentNullException(nameof(workbookReader));
            _modelBuilder = modelBuilder ?? throw new ArgumentNullException(nameof(modelBuilder));
            _substitutor = substitutor ?? throw new ArgumentNullException(nameof(substitutor));
            _layoutEngine = layoutEngine ?? throw new ArgumentNullException(nameof(layoutEngine));
            _renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));
            _ownedResources = ownedResources;
        }

        /// <summary>出力先パスを検証する(null・空・空白のみを弾く)。</summary>
        internal static void ValidateOutputPath(string outputPath)
        {
            if (outputPath is null)
            {
                throw new ArgumentNullException(nameof(outputPath));
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("出力先のパスが空です。", nameof(outputPath));
            }
        }

        internal static void ValidateMaxDigitWidth(double? maxDigitWidthPx)
        {
            if (maxDigitWidthPx is { } mdw && (mdw <= 0 || double.IsNaN(mdw) || double.IsInfinity(mdw)))
            {
                throw new ArgumentOutOfRangeException(nameof(maxDigitWidthPx), mdw, "最大数字幅は正の数である必要があります。");
            }
        }

        /// <summary>
        /// PDFを描画して <paramref name="output"/> へ書き出す。
        /// </summary>
        /// <remarks>
        /// 描画結果を出力先へ書き込むときの <see cref="IOException"/>(ディスクの空き不足、ネットワーク切断など)は、
        /// <see cref="IPdfRenderer"/> の実装によらず <see cref="PdfRenderingException"/> に包み、
        /// ファサードの例外を <see cref="UtsushiException"/> 階層にそろえる(要件6.4)。
        /// </remarks>
        internal void RenderToStream(PagedLayout layout, Stream output)
        {
            try
            {
                _renderer.Render(layout, output);
            }
            catch (FontNotAvailableException ex) when (LacksContext(ex))
            {
                throw WithContext(ex, ProcessingStage.Rendering, layout.DocumentName, layout.SheetName);
            }
            catch (IOException ex)
            {
                throw new PdfRenderingException(
                    $"PDFを出力先へ書き込めませんでした: {ex.Message}", layout.DocumentName, layout.SheetName, ex);
            }
        }

        /// <summary>PDFを描画してファイルへ書き出す。失敗時に不完全なファイルを残さない(要件5.4)。</summary>
        internal void RenderToFile(PagedLayout layout, string outputPath)
        {
            try
            {
                if (_renderer is SkiaPdfRenderer skia)
                {
                    skia.RenderToFile(layout, outputPath);
                    return;
                }

                using var buffer = new MemoryStream();
                _renderer.Render(layout, buffer);

                buffer.Position = 0;
                AtomicFileWriter.Write(outputPath, buffer, layout.DocumentName, layout.SheetName);
            }
            catch (FontNotAvailableException ex) when (LacksContext(ex))
            {
                throw WithContext(ex, ProcessingStage.Rendering, layout.DocumentName, layout.SheetName);
            }
        }

        /// <summary>
        /// フォントの解決(<see cref="FontResolver"/>)は文書を知らないため、文書名・シート名を持たない
        /// <see cref="FontNotAvailableException"/> を投げる。ファサードで補う必要があるかどうか。
        /// </summary>
        private static bool LacksContext(FontNotAvailableException ex) => ex.DocumentName is null || ex.SheetName is null;

        /// <summary>
        /// 文書名・シート名・処理段階を補った <see cref="FontNotAvailableException"/> で包み直す。
        /// 例外の型は変えず、元の例外は <see cref="Exception.InnerException"/> に入れる。
        /// </summary>
        private static FontNotAvailableException WithContext(
            FontNotAvailableException ex, ProcessingStage stage, string documentName, string sheetName) =>
            new(ex.FontName, ex.Message, stage, ex.DocumentName ?? documentName, ex.SheetName ?? sheetName, ex);

        /// <summary>
        /// PDF描画の手前まで(Parsing → 変換設定 → Substitution → Layout)を実行する(要件12)。
        /// </summary>
        /// <param name="xlsxStream">入力の Excel ファイル。</param>
        /// <param name="cellOverrides">セル番地 → 文字列(<see cref="Excel2Pdf.SetText"/>)。</param>
        /// <param name="documentName">文書名。</param>
        /// <param name="maxDigitWidthPx">最大数字幅。null ならブックの標準フォントから見積もる。</param>
        /// <param name="numericOverrides">セル番地 → 数値(<see cref="Excel2Pdf.SetValue(string, double)"/>)。</param>
        /// <param name="checkFit">文字の収まりと書式の近似を確かめるかどうか(診断のときだけ)。</param>
        /// <param name="sheetName">変換するシート。null ならアクティブシート。</param>
        /// <param name="onWorkbookRead">
        /// 読み込みの直後(書き込み・レイアウト計算の前)に呼ぶ。保存では未対応要素の扱いの判定に、診断では一覧の受け取りに使う
        /// (excel2pdf-consolidation 要件2.2・4)。
        /// </param>
        internal PagedLayout ComputeLayout(
            Stream xlsxStream,
            IReadOnlyDictionary<string, string>? cellOverrides,
            string? documentName,
            double? maxDigitWidthPx,
            IReadOnlyDictionary<string, double>? numericOverrides,
            bool checkFit,
            string? sheetName = null,
            Action<WorkbookModel>? onWorkbookRead = null)
        {
            ValidateMaxDigitWidth(maxDigitWidthPx);

            if (xlsxStream is null)
            {
                throw new ArgumentNullException(nameof(xlsxStream));
            }

            var name = documentName ?? string.Empty;
            var readOptions = new WorkbookReadOptions(
                name,
                SheetNameFilter: sheetName,
                ActiveSheetOnly: sheetName is null);

            var workbook = _workbookReader.Read(xlsxStream, readOptions);
            if (workbook.Sheets.Count != 1 || (sheetName is not null && workbook.Sheets[0].Name != sheetName))
            {
                // IWorkbookReader の契約(ActiveSheetOnly・SheetNameFilter なら対象の1枚だけ返す)に反する実装が
                // 差し替えられた場合に、対象でないシートを黙って変換しないよう止める。
                throw new InvalidOperationException(
                    $"{nameof(IWorkbookReader)} が変換対象の1枚(" + (sheetName is null ? "アクティブシート" : $"シート '{sheetName}'")
                    + $")ではなく {workbook.Sheets.Count} 枚のシートを返しました。");
            }

            var settings = ConversionSettings.Create(
                name,
                workbook.Sheets[0].Name,
                maxDigitWidthPx ?? ConversionSettings.EstimateMaxDigitWidthPx(workbook.DefaultFont.Name, workbook.DefaultFont.SizePt));

            onWorkbookRead?.Invoke(workbook);
            return BuildLayout(workbook, settings, cellOverrides, numericOverrides, checkFit);
        }

        private PagedLayout BuildLayout(
            WorkbookModel workbook,
            ConversionSettings settings,
            IReadOnlyDictionary<string, string>? cellOverrides,
            IReadOnlyDictionary<string, double>? numericOverrides,
            bool checkFit)
        {
            var report = _modelBuilder.Build(workbook, settings);

            if (cellOverrides is { Count: > 0 })
            {
                report = _substitutor.ApplyCellOverrides(report, cellOverrides);
            }

            if (numericOverrides is { Count: > 0 })
            {
                report = _substitutor.ApplyNumericCellOverrides(report, numericOverrides);
            }

            try
            {
                return _layoutEngine.Compute(report, checkFit);
            }
            catch (FontNotAvailableException ex) when (LacksContext(ex))
            {
                // レイアウト計算中の文字幅の計測(SkiaFontMetricsProvider)で起きた場合は、Stage=Layout として伝える。
                throw WithContext(ex, ProcessingStage.Layout, settings.DocumentName, settings.SheetName);
            }
        }

        /// <summary>
        /// 入力Excelファイルを開く。ファイルが無い/開けない場合も <see cref="UtsushiException"/> 階層(Stage=Parsing)へ統一する(要件6.4, 6.5)。
        /// </summary>
        internal static FileStream OpenInputFile(string xlsxPath, string? documentName)
        {
            try
            {
                return File.OpenRead(xlsxPath);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                // NotSupportedException は、パス文字列自体の形式が不正な場合に File.OpenRead が投げる。
                // ArgumentException は呼び出し側の契約違反を表すため、ここでは変換せずそのまま伝播させる。
                throw new InvalidExcelFileException(
                    $"入力ファイルを開けません: {xlsxPath}", InvalidExcelFileReason.Unknown, documentName, ex);
            }
        }

        public void Dispose() => _ownedResources?.Dispose();
    }
}
