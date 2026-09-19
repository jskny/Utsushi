using System;
using System.Collections.Generic;
using System.IO;
using Utsushi.Core.Exceptions;
using Utsushi.Layout;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi;

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
    /// 変換を実行し、PDFを <paramref name="output"/> へ書き出す。
    /// </summary>
    /// <param name="reportCode">帳票コード。</param>
    /// <param name="xlsxStream">入力となるExcelテンプレートのストリーム。</param>
    /// <param name="values">置換キー → 置換後の文字列。</param>
    /// <param name="output">PDFの出力先。</param>
    /// <exception cref="UtsushiException">
    /// 帳票定義・入力ファイル・置換値・レイアウト・描画のいずれかで失敗した場合。
    /// どの段階で失敗したかは <see cref="UtsushiException.Stage"/> で判別できる(要件6.4)。
    /// </exception>
    public void Convert(
        string reportCode,
        Stream xlsxStream,
        IReadOnlyDictionary<string, string> values,
        Stream output)
    {
        var layout = ComputeLayout(reportCode, xlsxStream, values);
        _renderer.Render(layout, output);
    }

    /// <summary>変換を実行し、PDFをファイルへ書き出す(失敗時に不完全なファイルを残さない)。</summary>
    public void ConvertToFile(
        string reportCode,
        string xlsxPath,
        IReadOnlyDictionary<string, string> values,
        string outputPath)
    {
        using var input = File.OpenRead(xlsxPath);
        var layout = ComputeLayout(reportCode, input, values);

        if (_renderer is SkiaPdfRenderer skia)
        {
            skia.RenderToFile(layout, outputPath);
            return;
        }

        using var buffer = new MemoryStream();
        _renderer.Render(layout, buffer);

        var directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory!);
        }

        using var file = File.Create(outputPath);
        buffer.Position = 0;
        buffer.CopyTo(file);
    }

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
        IReadOnlyDictionary<string, string> values)
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

        var definition = _definitions.Load(reportCode);

        var readOptions = new WorkbookReadOptions(
            definition.UnsupportedElements == UnsupportedElementPolicy.Error
                ? UnsupportedElementBehavior.Error
                : UnsupportedElementBehavior.Ignore,
            definition.ReportCode,
            definition.SheetName);

        var workbook = _workbookReader.Read(xlsxStream, readOptions);
        var report = _modelBuilder.Build(workbook, definition);
        var substituted = _substitutor.Apply(report, values);
        return _layoutEngine.Compute(substituted);
    }

    public void Dispose() => _ownedResources?.Dispose();
}
