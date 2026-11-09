using System;
using Utsushi.Layout;
using Utsushi.Parsing;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;

namespace Utsushi
{
    /// <summary>
    /// フォントの解決と描画の準備を持ち、複数の <see cref="Excel2Pdf"/> で共有するエンジン(excel2pdf-consolidation 要件1.2)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 大量に発行する場合やサーバー(ASP.NET Core など)で使う場合に、1つ作って共有する。<see cref="Excel2Pdf"/> を作るたびに
    /// フォントの解決をやり直さずに済む。スレッドセーフであり、複数のスレッドの <see cref="Excel2Pdf"/> から同時に使える。
    /// </para>
    /// <code>
    /// using var engine = Excel2PdfEngine.Create();
    /// using var pdf = new Excel2Pdf("template.xlsx", engine);
    /// pdf.SetText("B3", "株式会社サンプル 御中").Save("out.pdf");
    /// </code>
    /// <para>
    /// エンジンを渡して作った <see cref="Excel2Pdf"/> を破棄しても、エンジンは解放されない。使い終わったら呼び出し元で <see cref="Dispose"/> する。
    /// </para>
    /// </remarks>
    public sealed class Excel2PdfEngine : IDisposable
    {
        private readonly bool _ownsResources;
        private bool _disposed;

        /// <summary>部品を差し替えて作る(テストで決定的なフォントメトリクス・固定時刻を使うため。design.md 1.1)。</summary>
        /// <param name="fontResolver">フォントの解決。フォントの代替の問い合わせ(<see cref="Excel2Pdf.Check"/>)に使う。</param>
        /// <param name="ownsResources">true なら <see cref="Dispose"/> で <paramref name="fontResolver"/> を解放する。</param>
        internal Excel2PdfEngine(
            IWorkbookReader workbookReader,
            IReportModelBuilder modelBuilder,
            Substitution.ICellSubstitutor substitutor,
            IReportLayoutEngine layoutEngine,
            IPdfRenderer renderer,
            FontResolver fontResolver,
            bool ownsResources = true)
        {
            FontResolver = fontResolver ?? throw new ArgumentNullException(nameof(fontResolver));
            Converter = new ReportPdfConverter(
                workbookReader,
                new ReportPdfConverter.NoReportDefinitionRepository(),
                modelBuilder,
                substitutor,
                layoutEngine,
                renderer,
                ownsResources ? fontResolver : null);
            _ownsResources = ownsResources;
        }

        /// <summary>変換のパイプライン(帳票定義なしの経路だけを使う)。</summary>
        internal ReportPdfConverter Converter { get; }

        /// <summary>フォントの解決(フォントの代替の問い合わせに使う)。</summary>
        internal FontResolver FontResolver { get; }

        /// <summary>
        /// 既定の構成(OpenXml解析 + SkiaSharp描画)でエンジンを作る。
        /// </summary>
        /// <param name="fonts">フォント解決の設定。null の場合は厳格モード(無いフォントは <see cref="Core.Exceptions.FontNotAvailableException"/>)。</param>
        /// <param name="render">PDF出力の設定。null の場合は既定(フォント埋め込み)。</param>
        public static Excel2PdfEngine Create(FontResolverOptions? fonts = null, PdfRenderOptions? render = null)
        {
            var fontResolver = new FontResolver(fonts);
            try
            {
                var fontMetrics = new SkiaFontMetricsProvider(fontResolver);
                return new Excel2PdfEngine(
                    new Parsing.OpenXml.OpenXmlWorkbookReader(),
                    new ReportModelBuilder(),
                    new Substitution.CellSubstitutor(),
                    new ReportLayoutEngine(fontMetrics),
                    new SkiaPdfRenderer(fontMetrics, render),
                    fontResolver);
            }
            catch
            {
                fontResolver.Dispose();
                throw;
            }
        }

        /// <summary>破棄済みなら <see cref="ObjectDisposedException"/> を送出する(破棄したエンジンを渡した誤りを作成の時点で知らせる)。</summary>
        internal void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(Excel2PdfEngine));
            }
        }

        /// <summary>フォント解決のリソースを解放する。</summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            if (_ownsResources)
            {
                Converter.Dispose();
            }
        }
    }
}
