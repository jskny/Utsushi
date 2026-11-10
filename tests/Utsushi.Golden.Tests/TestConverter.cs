using System;
using System.Collections.Generic;
using System.IO;
using Utsushi.Layout;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.OpenXml;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.Substitution;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// 決定的なフォントメトリクス・固定時刻のエンジンで、<see cref="Excel2Pdf"/> を使って変換するテスト用の入口。
    /// </summary>
    /// <remarks>
    /// 以前の <c>ReportPdfConverter</c> の帳票定義なしの経路(未対応要素を無視して変換する)と同じ使い方を、
    /// <see cref="Excel2Pdf"/> の上で提供する(excel2pdf-consolidation タスク8.2)。未対応要素は無視する。
    /// </remarks>
    internal sealed class TestConverter : IDisposable
    {
        internal static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

        private readonly Excel2PdfEngine _engine;

        /// <param name="renderer">描画。null なら <see cref="SkiaPdfRenderer"/>(同梱の日本語フォントへの代替)。</param>
        public TestConverter(IPdfRenderer? renderer = null)
        {
            var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            _engine = new Excel2PdfEngine(
                new OpenXmlWorkbookReader(),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                renderer ?? new SkiaPdfRenderer(new SkiaFontMetricsProvider(fontResolver)),
                fontResolver);
        }

        /// <summary>レイアウト(文字の収まりの確認を含む)を求める。</summary>
        public PagedLayout ComputeLayout(
            Stream xlsxStream,
            IReadOnlyDictionary<string, string>? cellOverrides = null,
            string? documentName = null,
            double? maxDigitWidthPx = null)
        {
            using var pdf = Open(xlsxStream, cellOverrides, documentName, maxDigitWidthPx);
            return pdf.ComputeLayout(checkFit: true);
        }

        /// <summary>PDF をストリームへ書き出す。</summary>
        public void Convert(
            Stream xlsxStream,
            Stream output,
            IReadOnlyDictionary<string, string>? cellOverrides = null,
            string? documentName = null,
            double? maxDigitWidthPx = null)
        {
            using var pdf = Open(xlsxStream, cellOverrides, documentName, maxDigitWidthPx);
            pdf.Save(output);
        }

        /// <summary>PDF をファイルへ書き出す。文書名を省略すると入力のファイル名から拡張子を除いたもの。</summary>
        public void ConvertFile(
            string xlsxPath,
            string outputPath,
            IReadOnlyDictionary<string, string>? cellOverrides = null,
            string? documentName = null,
            double? maxDigitWidthPx = null)
        {
            using var pdf = new Excel2Pdf(xlsxPath, _engine, Options(documentName, maxDigitWidthPx));
            SetTexts(pdf, cellOverrides);
            pdf.Save(outputPath);
        }

        public void Dispose() => _engine.Dispose();

        private Excel2Pdf Open(
            Stream xlsxStream, IReadOnlyDictionary<string, string>? cellOverrides, string? documentName, double? maxDigitWidthPx)
        {
            var pdf = new Excel2Pdf(xlsxStream, _engine, Options(documentName, maxDigitWidthPx));
            SetTexts(pdf, cellOverrides);
            return pdf;
        }

        private static Excel2PdfOptions Options(string? documentName, double? maxDigitWidthPx) => new()
        {
            DocumentName = documentName,
            MaxDigitWidthPx = maxDigitWidthPx,
            UnsupportedElements = UnsupportedElementPolicy.Ignore,
        };

        private static void SetTexts(Excel2Pdf pdf, IReadOnlyDictionary<string, string>? cellOverrides)
        {
            foreach (var (cell, text) in cellOverrides ?? new Dictionary<string, string>())
            {
                pdf.SetText(cell, text);
            }
        }
    }
}
