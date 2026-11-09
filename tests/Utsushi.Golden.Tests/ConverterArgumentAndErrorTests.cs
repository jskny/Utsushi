using System;
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
using Utsushi.Substitution;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// ファサード(<see cref="Excel2Pdf"/>)の引数検証の順序と、下位レイヤーの例外を
    /// <see cref="UtsushiException"/> 階層へそろえる変換を確かめる(要件6.4)。
    /// </summary>
    public sealed class ConverterArgumentAndErrorTests
    {
        private static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

        private static Excel2PdfEngine CreateEngine(IFontMetricsProvider? layoutMetrics = null, IPdfRenderer? renderer = null)
        {
            var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            return new Excel2PdfEngine(
                new OpenXmlWorkbookReader(),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(layoutMetrics ?? new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                renderer ?? new SkiaPdfRenderer(new SkiaFontMetricsProvider(fontResolver)),
                fontResolver);
        }

        private static Excel2Pdf OpenInvoice(Excel2PdfEngine engine, string? path = null) =>
            new Excel2Pdf(path ?? TestPaths.SampleTemplate("invoice"), engine, new Excel2PdfOptions { DocumentName = "invoice" })
                .SetText("A3", "株式会社テスト製作所 御中");

        // -- L5: 引数の検証は計算を始める前に行う -------------------------------------------------

        [Fact]
        public void 保存先のストリームがnullなら入力を読む前にArgumentNullExceptionになる()
        {
            using var engine = CreateEngine();

            // 入力ファイルは存在しないが、先に出力先の検証で失敗する。
            using var pdf = OpenInvoice(engine, "no-such-input.xlsx");

            var ex = Assert.Throws<ArgumentNullException>(() => pdf.Save((Stream)null!));

            Assert.Equal("output", ex.ParamName);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void 保存先のパスが空なら入力を読む前に引数エラーになる(string? outputPath)
        {
            using var engine = CreateEngine();

            // 入力ファイルは存在しないが、先に出力先の検証で失敗する。
            using var pdf = OpenInvoice(engine, "no-such-input.xlsx");

            var ex = Assert.ThrowsAny<ArgumentException>(() => pdf.Save(outputPath!));

            Assert.Equal("outputPath", ex.ParamName);
            if (outputPath is null)
            {
                Assert.IsType<ArgumentNullException>(ex);
            }
        }

        // -- M1: 出力先への書き込みの失敗は PdfRenderingException にする --------------------------

        [Fact]
        public void 出力ストリームへの書き込みのIOExceptionはPdfRenderingExceptionになる()
        {
            using var engine = CreateEngine();
            using var pdf = OpenInvoice(engine);
            using var output = new FailingWriteStream();

            var ex = Assert.Throws<PdfRenderingException>(() => pdf.Save(output));

            Assert.IsAssignableFrom<IOException>(ex.InnerException);
            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
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
                using var engine = CreateEngine();
                using var pdf = OpenInvoice(engine);

                var ex = Assert.Throws<PdfRenderingException>(() => pdf.Save(Path.Combine(existingFile, "out.pdf")));

                Assert.IsAssignableFrom<IOException>(ex.InnerException);
                Assert.Equal("invoice", ex.ReportCode);
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        // -- L1: FontNotAvailableException に文書名・シート名・処理段階を補う -------------------

        [Fact]
        public void レイアウト計算中にフォントが見つからなければStageLayoutで文書名とシート名が入る()
        {
            using var engine = CreateEngine(layoutMetrics: new MissingFontMetricsProvider());
            using var pdf = OpenInvoice(engine);

            var ex = Assert.Throws<FontNotAvailableException>(() => pdf.ToPdfBytes());

            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
            Assert.Equal(MissingFontMetricsProvider.FontName, ex.FontName);
            var inner = Assert.IsType<FontNotAvailableException>(ex.InnerException);
            Assert.Null(inner.ReportCode);
        }

        [Fact]
        public void 描画中にフォントが見つからなければStageRenderingで文書名とシート名が入る()
        {
            using var engine = CreateEngine(renderer: new MissingFontRenderer());
            using var pdf = OpenInvoice(engine);

            var ex = Assert.Throws<FontNotAvailableException>(() => pdf.ToPdfBytes());

            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
            Assert.IsType<FontNotAvailableException>(ex.InnerException);
        }

        [Fact]
        public void ファイル出力の描画中にフォントが見つからなくても文書名とシート名が入る()
        {
            using var engine = CreateEngine(renderer: new MissingFontRenderer());
            using var pdf = OpenInvoice(engine);
            var outputPath = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N") + ".pdf");

            var ex = Assert.Throws<FontNotAvailableException>(() => pdf.Save(outputPath));

            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal("請求書", ex.SheetName);
            Assert.False(File.Exists(outputPath));
        }

        /// <summary>書き込みのたびに IOException を投げる(ディスクの空き不足・ネットワーク切断の代わり)。</summary>
        private sealed class FailingWriteStream : MemoryStream
        {
            public override void Write(byte[] buffer, int offset, int count) =>
                throw new IOException("ディスクに空きがありません(テスト)。");

            public override void Write(ReadOnlySpan<byte> buffer) =>
                throw new IOException("ディスクに空きがありません(テスト)。");
        }

        /// <summary>FontResolver と同様に、文書名・シート名を持たない FontNotAvailableException を投げる。</summary>
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
