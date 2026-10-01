using System;
using System.Collections.Generic;
using System.IO;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing;
using Utsushi.Parsing.Model;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi
{
    /// <summary>
    /// 5つのレイヤーを組み立て、「帳票コード + 置換値 + 入力Excel」から「PDF」を得るユースケース。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 呼び出し元プロダクトはこのクラスだけを参照すればよい。
    /// 各レイヤーの実装は差し替え可能なようにコンストラクタで受け取る。
    /// </para>
    /// <para>
    /// 既定の構成は <see cref="CreateDefault"/> で得られる。
    /// このクラスは <see cref="FontResolver"/> を保持するため、使い終わったら <see cref="Dispose"/> すること。
    /// </para>
    /// </remarks>
    public sealed class ReportPdfConverter : IDisposable
    {
        private readonly IWorkbookReader _workbookReader;
        private readonly IReportDefinitionRepository _definitions;
        private readonly IReportModelBuilder _modelBuilder;
        private readonly Substitution.ICellSubstitutor _substitutor;
        private readonly IReportLayoutEngine _layoutEngine;
        private readonly IPdfRenderer _renderer;
        private readonly IDisposable? _ownedResources;

        public ReportPdfConverter(
            IWorkbookReader workbookReader,
            IReportDefinitionRepository definitions,
            IReportModelBuilder modelBuilder,
            Substitution.ICellSubstitutor substitutor,
            IReportLayoutEngine layoutEngine,
            IPdfRenderer renderer,
            IDisposable? ownedResources = null)
        {
            _workbookReader = workbookReader ?? throw new ArgumentNullException(nameof(workbookReader));
            _definitions = definitions ?? throw new ArgumentNullException(nameof(definitions));
            _modelBuilder = modelBuilder ?? throw new ArgumentNullException(nameof(modelBuilder));
            _substitutor = substitutor ?? throw new ArgumentNullException(nameof(substitutor));
            _layoutEngine = layoutEngine ?? throw new ArgumentNullException(nameof(layoutEngine));
            _renderer = renderer ?? throw new ArgumentNullException(nameof(renderer));
            _ownedResources = ownedResources;
        }

        /// <summary>
        /// 既定の構成(OpenXml解析 + SkiaSharp描画)でコンバータを組み立てる。
        /// </summary>
        /// <param name="reportDefinitionRoot">帳票定義のルートディレクトリ。</param>
        /// <param name="fontOptions">フォント解決の設定。null の場合は厳格モード。</param>
        /// <param name="renderOptions">PDF出力の設定。null の場合は既定(フォント埋め込み)。</param>
        public static ReportPdfConverter CreateDefault(
            string reportDefinitionRoot,
            FontResolverOptions? fontOptions = null,
            PdfRenderOptions? renderOptions = null)
        {
            var fontResolver = new FontResolver(fontOptions);
            try
            {
                var fontMetrics = new SkiaFontMetricsProvider(fontResolver);
                return new ReportPdfConverter(
                    new Utsushi.Parsing.OpenXml.OpenXmlWorkbookReader(),
                    new FileSystemReportDefinitionRepository(reportDefinitionRoot),
                    new ReportModelBuilder(),
                    new Substitution.CellSubstitutor(),
                    new ReportLayoutEngine(fontMetrics),
                    new SkiaPdfRenderer(fontMetrics, renderOptions),
                    fontResolver);
            }
            catch
            {
                fontResolver.Dispose();
                throw;
            }
        }

        /// <summary>
        /// 帳票定義を使わない構成(OpenXml解析 + SkiaSharp描画)でコンバータを組み立てる(要件12)。
        /// </summary>
        /// <remarks>
        /// 帳票定義なしの変換(<see cref="ConvertWithoutDefinition"/> 等)だけを使う呼び出し元向け。
        /// この構成で帳票コードを指定して変換すると <see cref="ReportDefinitionNotFoundException"/> になる。
        /// </remarks>
        /// <param name="fontOptions">フォント解決の設定。null の場合は厳格モード。</param>
        /// <param name="renderOptions">PDF出力の設定。null の場合は既定(フォント埋め込み)。</param>
        public static ReportPdfConverter CreateDefault(
            FontResolverOptions? fontOptions = null,
            PdfRenderOptions? renderOptions = null)
        {
            var fontResolver = new FontResolver(fontOptions);
            try
            {
                var fontMetrics = new SkiaFontMetricsProvider(fontResolver);
                return new ReportPdfConverter(
                    new Utsushi.Parsing.OpenXml.OpenXmlWorkbookReader(),
                    new NoReportDefinitionRepository(),
                    new ReportModelBuilder(),
                    new Substitution.CellSubstitutor(),
                    new ReportLayoutEngine(fontMetrics),
                    new SkiaPdfRenderer(fontMetrics, renderOptions),
                    fontResolver);
            }
            catch
            {
                fontResolver.Dispose();
                throw;
            }
        }

        /// <summary>
        /// 変換を実行し、PDFを <paramref name="output"/> へ書き出す。
        /// </summary>
        /// <param name="reportCode">帳票コード。</param>
        /// <param name="xlsxStream">入力となるExcelテンプレートのストリーム。</param>
        /// <param name="values">置換キー → 置換後の文字列。</param>
        /// <param name="output">PDFの出力先。</param>
        /// <param name="cellOverrides">
        /// セル番地(A1形式) → 上書き後の文字列。帳票定義への登録有無に関わらず直接指定できる(要件2.7)。
        /// 指定不要な場合は null または空でよい。
        /// </param>
        /// <exception cref="UtsushiException">
        /// 帳票定義・入力ファイル・置換値・レイアウト・描画のいずれかで失敗した場合。
        /// どの段階で失敗したかは <see cref="UtsushiException.Stage"/> で判別できる(要件6.4)。
        /// </exception>
        public void Convert(
            string reportCode,
            Stream xlsxStream,
            IReadOnlyDictionary<string, string> values,
            Stream output,
            IReadOnlyDictionary<string, string>? cellOverrides = null)
        {
            // 変換(解析・レイアウト計算)を始める前に、すべての引数を検証する。
            ValidateConvertArguments(reportCode, xlsxStream, values);
            if (output is null)
            {
                throw new ArgumentNullException(nameof(output));
            }

            var layout = ComputeLayout(reportCode, xlsxStream, values, cellOverrides);
            RenderToStream(layout, output);
        }

        /// <summary>変換を実行し、PDFをファイルへ書き出す(失敗時に不完全なファイルを残さない)。</summary>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="reportCode"/>・<paramref name="xlsxPath"/>・<paramref name="values"/>・<paramref name="outputPath"/> が null の場合。
        /// </exception>
        /// <exception cref="ArgumentException"><paramref name="outputPath"/> が空または空白のみの場合。</exception>
        /// <exception cref="UtsushiException">
        /// 帳票定義・入力ファイル・置換値・レイアウト・描画・出力ファイルの書き込みのいずれかで失敗した場合。
        /// </exception>
        public void ConvertToFile(
            string reportCode,
            string xlsxPath,
            IReadOnlyDictionary<string, string> values,
            string outputPath,
            IReadOnlyDictionary<string, string>? cellOverrides = null)
        {
            // 変換(入力ファイルの読み込み・レイアウト計算)を始める前に、すべての引数を検証する。
            if (reportCode is null)
            {
                throw new ArgumentNullException(nameof(reportCode));
            }

            if (xlsxPath is null)
            {
                throw new ArgumentNullException(nameof(xlsxPath));
            }

            if (values is null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            ValidateOutputPath(outputPath);

            using var input = OpenInputFile(xlsxPath, reportCode);
            var layout = ComputeLayout(reportCode, input, values, cellOverrides);
            RenderToFile(layout, outputPath);
        }

        /// <summary>
        /// 帳票定義を使わずに変換し、PDFを <paramref name="output"/> へ書き出す(要件12)。
        /// </summary>
        /// <remarks>
        /// ブックのアクティブシート(非表示なら表示されている最初のシート)1枚を、既定値の帳票定義
        /// (置換キーなし・サポート外要素は無視・印刷範囲はExcelの設定)で変換する。
        /// 見た目の一致を保証しないベストエフォートの変換である(<c>.kiro/steering/product.md</c>)。
        /// </remarks>
        /// <param name="xlsxStream">入力となるExcelファイルのストリーム。</param>
        /// <param name="output">PDFの出力先。</param>
        /// <param name="cellOverrides">セル番地(A1形式) → 上書き後の文字列(要件2.7)。不要なら null。</param>
        /// <param name="documentName">
        /// PDFのタイトル・ヘッダー/フッターのファイル名(<c>&amp;F</c>)・エラー情報に使う名前。null の場合は空文字列。
        /// </param>
        /// <param name="maxDigitWidthPx">
        /// 列幅の換算に使う最大数字幅(ピクセル)。null の場合はブックの標準フォントから見積もる
        /// (<see cref="ReportDefinition.EstimateMaxDigitWidthPx"/>)。列幅がExcelとずれる場合に指定する。
        /// </param>
        /// <exception cref="UtsushiException">入力ファイル・レイアウト・描画のいずれかで失敗した場合。</exception>
        public void ConvertWithoutDefinition(
            Stream xlsxStream,
            Stream output,
            IReadOnlyDictionary<string, string>? cellOverrides = null,
            string? documentName = null,
            double? maxDigitWidthPx = null)
        {
            if (xlsxStream is null)
            {
                throw new ArgumentNullException(nameof(xlsxStream));
            }

            if (output is null)
            {
                throw new ArgumentNullException(nameof(output));
            }

            ValidateMaxDigitWidth(maxDigitWidthPx);

            var layout = ComputeLayoutWithoutDefinition(xlsxStream, cellOverrides, documentName, maxDigitWidthPx);
            RenderToStream(layout, output);
        }

        /// <summary>
        /// 帳票定義を使わずに変換し、PDFをファイルへ書き出す(要件12。失敗時に不完全なファイルを残さない)。
        /// </summary>
        /// <param name="xlsxPath">入力Excelファイルのパス。</param>
        /// <param name="outputPath">PDFの出力先パス。</param>
        /// <param name="cellOverrides">セル番地(A1形式) → 上書き後の文字列(要件2.7)。不要なら null。</param>
        /// <param name="documentName">
        /// PDFのタイトル・ヘッダー/フッターのファイル名・エラー情報に使う名前。null の場合は入力ファイル名から
        /// 拡張子を除いたもの(要件12.5)。ファイル名に取引先名などを含み、PDFのメタデータやログに出したくない場合に指定する。
        /// </param>
        /// <param name="maxDigitWidthPx">列幅の換算に使う最大数字幅(ピクセル)。null の場合はブックの標準フォントから見積もる。</param>
        public void ConvertFileWithoutDefinition(
            string xlsxPath,
            string outputPath,
            IReadOnlyDictionary<string, string>? cellOverrides = null,
            string? documentName = null,
            double? maxDigitWidthPx = null)
        {
            // 変換(入力ファイルの読み込み・レイアウト計算)を始める前に、すべての引数を検証する。
            if (xlsxPath is null)
            {
                throw new ArgumentNullException(nameof(xlsxPath));
            }

            ValidateOutputPath(outputPath);
            ValidateMaxDigitWidth(maxDigitWidthPx);

            documentName ??= Path.GetFileNameWithoutExtension(xlsxPath);

            using var input = OpenInputFile(xlsxPath, documentName);
            var layout = ComputeLayoutWithoutDefinition(input, cellOverrides, documentName, maxDigitWidthPx);
            RenderToFile(layout, outputPath);
        }

        /// <summary>出力先パスを検証する(null・空・空白のみを弾く)。</summary>
        private static void ValidateOutputPath(string outputPath)
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

        private static void ValidateConvertArguments(
            string reportCode, Stream xlsxStream, IReadOnlyDictionary<string, string> values)
        {
            if (reportCode is null)
            {
                throw new ArgumentNullException(nameof(reportCode));
            }

            if (xlsxStream is null)
            {
                throw new ArgumentNullException(nameof(xlsxStream));
            }

            if (values is null)
            {
                throw new ArgumentNullException(nameof(values));
            }
        }

        private static void ValidateMaxDigitWidth(double? maxDigitWidthPx)
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
        private void RenderToStream(PagedLayout layout, Stream output)
        {
            try
            {
                _renderer.Render(layout, output);
            }
            catch (FontNotAvailableException ex) when (LacksContext(ex))
            {
                throw WithContext(ex, ProcessingStage.Rendering, layout.ReportCode, layout.SheetName);
            }
            catch (IOException ex)
            {
                throw new PdfRenderingException(
                    $"PDFを出力先へ書き込めませんでした: {ex.Message}", layout.ReportCode, layout.SheetName, ex);
            }
        }

        private void RenderToFile(PagedLayout layout, string outputPath)
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
                AtomicFileWriter.Write(outputPath, buffer, layout.ReportCode, layout.SheetName);
            }
            catch (FontNotAvailableException ex) when (LacksContext(ex))
            {
                throw WithContext(ex, ProcessingStage.Rendering, layout.ReportCode, layout.SheetName);
            }
        }

        /// <summary>
        /// フォントの解決(<see cref="FontResolver"/>)は帳票を知らないため、帳票コード・シート名を持たない
        /// <see cref="FontNotAvailableException"/> を投げる。ファサードで補う必要があるかどうか。
        /// </summary>
        private static bool LacksContext(FontNotAvailableException ex) => ex.ReportCode is null || ex.SheetName is null;

        /// <summary>
        /// 帳票コード・シート名・処理段階を補った <see cref="FontNotAvailableException"/> で包み直す。
        /// 例外の型は変えず、元の例外は <see cref="Exception.InnerException"/> に入れる。
        /// </summary>
        private static FontNotAvailableException WithContext(
            FontNotAvailableException ex, ProcessingStage stage, string reportCode, string sheetName) =>
            new(ex.FontName, ex.Message, stage, ex.ReportCode ?? reportCode, ex.SheetName ?? sheetName, ex);

        /// <summary>
        /// PDF描画の手前まで(Parsing → ReportDefinition → Substitution → Layout)を実行する。
        /// </summary>
        /// <remarks>
        /// ゴールデンテストは描画結果のバイナリではなくこのレイアウト結果を比較対象とするため、
        /// 公開メソッドとして切り出している(design.md「テスト戦略」)。
        /// </remarks>
        public PagedLayout ComputeLayout(
            string reportCode,
            Stream xlsxStream,
            IReadOnlyDictionary<string, string> values,
            IReadOnlyDictionary<string, string>? cellOverrides = null)
        {
            ValidateConvertArguments(reportCode, xlsxStream, values);

            var definition = _definitions.Load(reportCode);

            var readOptions = new WorkbookReadOptions(
                definition.UnsupportedElements == UnsupportedElementPolicy.Error
                    ? UnsupportedElementBehavior.Error
                    : UnsupportedElementBehavior.Ignore,
                definition.ReportCode,
                definition.SheetName);

            var workbook = ReadWorkbook(xlsxStream, readOptions, definition);
            return BuildLayout(workbook, definition, values, cellOverrides);
        }

        /// <summary>
        /// 帳票定義を使わずに、PDF描画の手前まで(Parsing → 既定値の帳票定義 → Substitution → Layout)を実行する(要件12)。
        /// </summary>
        public PagedLayout ComputeLayoutWithoutDefinition(
            Stream xlsxStream,
            IReadOnlyDictionary<string, string>? cellOverrides = null,
            string? documentName = null,
            double? maxDigitWidthPx = null)
        {
            ValidateMaxDigitWidth(maxDigitWidthPx);

            if (xlsxStream is null)
            {
                throw new ArgumentNullException(nameof(xlsxStream));
            }

            var name = documentName ?? string.Empty;
            var readOptions = new WorkbookReadOptions(
                UnsupportedElementBehavior.Ignore,
                name,
                SheetNameFilter: null,
                ActiveSheetOnly: true);

            var workbook = _workbookReader.Read(xlsxStream, readOptions);
            if (workbook.Sheets.Count != 1)
            {
                // IWorkbookReader の契約(ActiveSheetOnly なら1枚だけ返す)に反する実装が差し替えられた場合に、
                // アクティブでないシートを黙って変換しないよう止める。
                throw new InvalidOperationException(
                    $"{nameof(IWorkbookReader)} が {nameof(WorkbookReadOptions.ActiveSheetOnly)} を指定したのに "
                    + $"{workbook.Sheets.Count} 枚のシートを返しました。");
            }

            var definition = ReportDefinition.CreateWithoutDefinition(
                name,
                workbook.Sheets[0].Name,
                maxDigitWidthPx ?? ReportDefinition.EstimateMaxDigitWidthPx(workbook.DefaultFont.Name, workbook.DefaultFont.SizePt));

            return BuildLayout(workbook, definition, NoValues, cellOverrides);
        }

        private static readonly IReadOnlyDictionary<string, string> NoValues = new Dictionary<string, string>();

        private PagedLayout BuildLayout(
            WorkbookModel workbook,
            ReportDefinition definition,
            IReadOnlyDictionary<string, string> values,
            IReadOnlyDictionary<string, string>? cellOverrides)
        {
            var report = _modelBuilder.Build(workbook, definition);
            var substituted = _substitutor.Apply(report, values);

            if (cellOverrides is { Count: > 0 })
            {
                substituted = _substitutor.ApplyCellOverrides(substituted, cellOverrides);
            }

            try
            {
                return _layoutEngine.Compute(substituted);
            }
            catch (FontNotAvailableException ex) when (LacksContext(ex))
            {
                // レイアウト計算中の文字幅の計測(SkiaFontMetricsProvider)で起きた場合は、Stage=Layout として伝える。
                throw WithContext(ex, ProcessingStage.Layout, definition.ReportCode, definition.SheetName);
            }
        }

        /// <summary>
        /// ワークブックを読み取る。帳票定義が期待するシートが実在しない場合、Parsing層は
        /// 「読み込み対象を絞り込んだ結果0件だった」ことしか知らないため <see cref="InvalidExcelFileException"/>
        /// (Stage=Parsing)を投げるが、これは実際にはファイル自体の不正ではなく帳票定義と
        /// ワークブックの構造不一致(要件1.4)である。呼び出し元が <see cref="UtsushiException.Stage"/>
        /// で正しく分岐できるよう、ここで <see cref="ReportStructureMismatchException"/>(Stage=ReportDefinition)
        /// に読み替える。
        /// </summary>
        private WorkbookModel ReadWorkbook(Stream xlsxStream, WorkbookReadOptions readOptions, ReportDefinition definition)
        {
            try
            {
                return _workbookReader.Read(xlsxStream, readOptions);
            }
            catch (InvalidExcelFileException ex) when (ex.Reason == InvalidExcelFileReason.NoWorksheet
                && readOptions.SheetNameFilter is not null)
            {
                throw new ReportStructureMismatchException(
                    ex.Message, definition.ReportCode, definition.SheetName, innerException: ex);
            }
        }

        /// <summary>
        /// 入力Excelファイルを開く。ファイルが無い/開けない場合も、<see cref="Convert"/>(ストリーム版)と
        /// 同様に <see cref="UtsushiException"/> 階層(Stage=Parsing)へ統一する(要件6.4, 6.5)。
        /// </summary>
        private static FileStream OpenInputFile(string xlsxPath, string? reportCode)
        {
            try
            {
                return File.OpenRead(xlsxPath);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                // NotSupportedExceptionは、パス文字列自体の形式が不正な場合(例: コロンを含む等)に
                // File.OpenReadが投げる。ArgumentException/ArgumentNullExceptionは、値が渡されなかった/
                // 空という呼び出し側の契約違反を表すため、あえてここでは変換せずそのまま伝播させる
                // (ComputeLayoutの引数nullチェックと同様の扱い)。
                throw new InvalidExcelFileException(
                    $"入力ファイルを開けません: {xlsxPath}", InvalidExcelFileReason.Unknown, reportCode, ex);
            }
        }

        public void Dispose() => _ownedResources?.Dispose();

        /// <summary>
        /// 帳票定義ルートを持たない構成(<see cref="CreateDefault(FontResolverOptions?, PdfRenderOptions?)"/>)で使う。
        /// 帳票コードを指定した変換は常に「帳票定義が見つからない」エラーにする(要件1.4)。
        /// </summary>
        private sealed class NoReportDefinitionRepository : IReportDefinitionRepository
        {
            public ReportDefinition Load(string reportCode) =>
                throw new ReportDefinitionNotFoundException(
                    reportCode,
                    $"帳票コード '{reportCode}' の帳票定義が見つかりません"
                    + "(帳票定義のルートディレクトリを指定せずに作成したコンバータです)。");

            public IReadOnlyCollection<string> ListReportCodes() => Array.Empty<string>();
        }
    }
}
