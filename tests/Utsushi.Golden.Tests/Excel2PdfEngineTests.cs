using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Utsushi.Layout.Model;
using Utsushi.Rendering;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// 複数の <see cref="Excel2Pdf"/> で共有するエンジン(excel2pdf-consolidation 要件1.2)。請求書サンプルを使う。
    /// </summary>
    public sealed class Excel2PdfEngineTests
    {
        private static readonly string Invoice = TestPaths.SampleTemplate("invoice");

        [Fact]
        public void エンジンを渡して作ったExcel2Pdfを破棄してもエンジンは使い続けられる()
        {
            using var engine = Excel2PdfEngine.Create(FontResolverOptions.AllowFallback());

            using (var pdf = new Excel2Pdf(Invoice, engine))
            {
                pdf.SetText("A3", "一通目");
                Assert.True(pdf.ToPdfBytes().Length > 0);
            }

            using var bytes = new MemoryStream(File.ReadAllBytes(Invoice));
            using var next = new Excel2Pdf(bytes, engine);
            next.SetText("A3", "二通目");
            Assert.True(next.ToPdfBytes().Length > 0);
        }

        [Fact]
        public void エンジンを渡す場合はフォントとPDF出力の設定を指定できない()
        {
            using var engine = Excel2PdfEngine.Create(FontResolverOptions.AllowFallback());

            Assert.Throws<ArgumentException>(
                () => new Excel2Pdf(Invoice, engine, new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback() }));
            Assert.Throws<ArgumentException>(
                () => new Excel2Pdf(File.ReadAllBytes(Invoice), engine, new Excel2PdfOptions { Render = PdfRenderOptions.Default }));
        }

        [Fact]
        public void エンジンにnullを渡すと引数のエラーにする()
        {
            Assert.Throws<ArgumentNullException>(() => new Excel2Pdf(Invoice, (Excel2PdfEngine)null!));
        }

        [Fact]
        public void 共有したエンジンで同時に変換しても値が混ざらない()
        {
            using var engine = Excel2PdfEngine.Create(FontResolverOptions.AllowFallback());
            var failures = new ConcurrentBag<string>();

            Parallel.For(0, 40, new ParallelOptions { MaxDegreeOfParallelism = 8 }, i =>
            {
                var name = $"顧客{i:000} 御中";
                using var pdf = new Excel2Pdf(Invoice, engine);
                pdf.SetText("A3", name).SetValue("F12", i * 1000);

                Assert.True(pdf.ToPdfBytes().Length > 0);

                var texts = pdf.ComputeLayout(checkFit: false).Pages
                    .SelectMany(p => p.Commands.OfType<TextCommand>())
                    .Select(t => t.Text)
                    .ToList();
                var expectedAmount = (i * 1000).ToString("\\¥#,##0", System.Globalization.CultureInfo.InvariantCulture);
                if (!texts.Contains(name) || !texts.Contains(expectedAmount)
                    || texts.Any(t => t.StartsWith("顧客", StringComparison.Ordinal) && t != name))
                {
                    failures.Add(i.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            });

            Assert.Empty(failures);
        }

        [Fact]
        public void 破棄したエンジンを渡すと作成の時点でエラーにする()
        {
            var engine = Excel2PdfEngine.Create(FontResolverOptions.AllowFallback());
            engine.Dispose();

            Assert.Throws<ObjectDisposedException>(() => new Excel2Pdf(Invoice, engine));
        }

        [Fact]
        public void 破棄したエンジンは二度破棄しても例外にならない()
        {
            var engine = Excel2PdfEngine.Create(FontResolverOptions.AllowFallback());
            engine.Dispose();
            engine.Dispose();
        }
    }
}
