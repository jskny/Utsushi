using System;
using System.Collections.Generic;
using System.IO;
using Utsushi;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.ReportDefinitions.Model;
using Utsushi.Substitution;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// ファサード(<see cref="ReportPdfConverter"/>)の引数検証の順序と、下位レイヤーの例外を
    /// <see cref="UtsushiException"/> 階層へそろえる変換を確かめる(要件6.4)。
    /// </summary>
    public sealed class ConverterArgumentAndErrorTests
    {
        private static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

        private static Dictionary<string, string> InvoiceValues() => new()
        {
            ["CustomerName"] = "株式会社テスト製作所 御中",
            ["InvoiceNo"] = "INV-2026-0417",
            ["TotalAmount"] = "¥2,153,800",
        };

        private static ReportPdfConverter CreateConverter(
            FontResolver fontResolver,
            IReportDefinitionRepository? definitions = null,
            IFontMetricsProvider? layoutMetrics = null,
            IPdfRenderer? renderer = null) =>
            new(
                new OpenXmlWorkbookReader(),
                definitions ?? new FileSystemReportDefinitionRepository(TestPaths.SampleReportsRoot),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(layoutMetrics ?? new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                renderer ?? new SkiaPdfRenderer(new SkiaFontMetricsProvider(fontResolver)));

        // -- L5: 引数の検証は計算を始める前に行う -------------------------------------------------

        [Fact]
        public void Convertのoutputがnullなら帳票定義を読む前にArgumentNullExceptionになる()
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver, definitions: new MustNotBeCalledRepository());
            using var input = new MemoryStream();

            var ex = Assert.Throws<ArgumentNullException>(
                () => converter.Convert("invoice", input, InvoiceValues(), null!));

            Assert.Equal("output", ex.ParamName);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void ConvertToFileの出力先が空なら入力を読む前に引数エラーになる(string? outputPath)
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver, definitions: new MustNotBeCalledRepository());

            // 入力ファイルは存在しないが、先に出力先の検証で失敗する。
            var ex = Assert.ThrowsAny<ArgumentException>(
                () => converter.ConvertToFile("invoice", "no-such-input.xlsx", InvoiceValues(), outputPath!));

            Assert.Equal("outputPath", ex.ParamName);
            if (outputPath is null)
            {
                Assert.IsType<ArgumentNullException>(ex);
            }
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void ConvertFileWithoutDefinitionの出力先が空なら入力を読む前に引数エラーになる(string? outputPath)
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver);

            // 以前は入力ファイルを開いてから検証していたため、InvalidExcelFileException になっていた。
            var ex = Assert.ThrowsAny<ArgumentException>(
                () => converter.ConvertFileWithoutDefinition("no-such-input.xlsx", outputPath!));

            Assert.Equal("outputPath", ex.ParamName);
        }

        // -- M1: 出力先への書き込みの失敗は PdfRenderingException にする --------------------------

        [Fact]
        public void 出力ストリームへの書き込みのIOExceptionはPdfRenderingExceptionになる()
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var output = new FailingWriteStream();

            var ex = Assert.Throws<PdfRenderingException>(
                () => converter.Convert("invoice", input, InvoiceValues(), output));

            Assert.IsAssignableFrom<IOException>(ex.InnerException);
            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
        }

        [Fact]
        public void 定義なしの変換でも出力ストリームへの書き込みのIOExceptionはPdfRenderingExceptionになる()
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var output = new FailingWriteStream();

            var ex = Assert.Throws<PdfRenderingException>(
                () => converter.ConvertWithoutDefinition(input, output, documentName: "文書"));

            Assert.IsAssignableFrom<IOException>(ex.InnerException);
            Assert.Equal("文書", ex.ReportCode);
        }

        [Fact]
        public void 出力先の途中に既存のファイルがあるとPdfRenderingExceptionになる()
        {
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var existingFile = Path.Combine(directory, "existing.txt");
            File.WriteAllText(existingFile, "ファイル");

            try
            {
                using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
                using var converter = CreateConverter(fontResolver);

                var ex = Assert.Throws<PdfRenderingException>(
                    () => converter.ConvertToFile(
                        "invoice", TestPaths.SampleTemplate("invoice"), InvoiceValues(), Path.Combine(existingFile, "out.pdf")));

                Assert.IsAssignableFrom<IOException>(ex.InnerException);
                Assert.Equal("invoice", ex.ReportCode);
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        // -- L1: FontNotAvailableException に帳票コード・シート名・処理段階を補う -------------------

        [Fact]
        public void レイアウト計算中にフォントが見つからなければStageLayoutで帳票コードとシート名が入る()
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver, layoutMetrics: new MissingFontMetricsProvider());
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<FontNotAvailableException>(
                () => converter.ComputeLayout("invoice", input, InvoiceValues()));

            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
            Assert.Equal(MissingFontMetricsProvider.FontName, ex.FontName);
            var inner = Assert.IsType<FontNotAvailableException>(ex.InnerException);
            Assert.Null(inner.ReportCode);
        }

        [Fact]
        public void 描画中にフォントが見つからなければStageRenderingで帳票コードとシート名が入る()
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver, renderer: new MissingFontRenderer());
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var output = new MemoryStream();

            var ex = Assert.Throws<FontNotAvailableException>(
                () => converter.Convert("invoice", input, InvoiceValues(), output));

            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
            Assert.IsType<FontNotAvailableException>(ex.InnerException);
        }

        [Fact]
        public void ファイル出力の描画中にフォントが見つからなくても帳票コードとシート名が入る()
        {
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = CreateConverter(fontResolver, renderer: new MissingFontRenderer());
            var outputPath = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N") + ".pdf");

            var ex = Assert.Throws<FontNotAvailableException>(
                () => converter.ConvertToFile("invoice", TestPaths.SampleTemplate("invoice"), InvoiceValues(), outputPath));

            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
            Assert.False(File.Exists(outputPath));
        }

        private sealed class MustNotBeCalledRepository : IReportDefinitionRepository
        {
            public ReportDefinition Load(string reportCode) =>
                throw new InvalidOperationException("引数の検証より先に帳票定義を読んではいけない。");

            public IReadOnlyCollection<string> ListReportCodes() => Array.Empty<string>();
        }

        /// <summary>書き込みのたびに IOException を投げる(ディスクの空き不足・ネットワーク切断の代わり)。</summary>
        private sealed class FailingWriteStream : MemoryStream
        {
            public override void Write(byte[] buffer, int offset, int count) =>
                throw new IOException("ディスクに空きがありません(テスト)。");

            public override void Write(ReadOnlySpan<byte> buffer) =>
                throw new IOException("ディスクに空きがありません(テスト)。");
        }

        /// <summary>FontResolver と同様に、帳票コード・シート名を持たない FontNotAvailableException を投げる。</summary>
        private sealed class MissingFontMetricsProvider : IFontMetricsProvider
        {
            public const string FontName = "存在しないフォント-ZZZ";

            public FontMetrics GetMetrics(FontStyle font) => throw Missing();

            public double MeasureTextWidth(FontStyle font, string text) => throw Missing();

            private static FontNotAvailableException Missing() =>
                new(FontName, $"フォント '{FontName}' が実行環境に見つかりません(テスト)。");
        }

        private sealed class MissingFontRenderer : IPdfRenderer
        {
            public void Render(PagedLayout layout, Stream output) =>
                throw new FontNotAvailableException("存在しないフォント-ZZZ", "フォントが見つかりません(テスト)。");
        }
    }
}
