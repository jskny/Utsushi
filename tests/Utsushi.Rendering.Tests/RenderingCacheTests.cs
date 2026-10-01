using System;
using System.Collections.Generic;
using System.IO;
using SkiaSharp;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Utsushi.Rendering.Fonts;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// 計測・描画のキャッシュ(フォントメトリクス・SKFont・破線パターン・デコード済み画像)が、
    /// 結果を変えずに使い回されることを確かめる。
    /// </summary>
    public sealed class RenderingCacheTests : IDisposable
    {
        private readonly FontResolver _fontResolver;
        private readonly SkiaFontMetricsProvider _metrics;

        public RenderingCacheTests()
        {
            _fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            _metrics = new SkiaFontMetricsProvider(_fontResolver);
        }

        public void Dispose() => _fontResolver.Dispose();

        [Fact]
        public void フォントメトリクスは記録した値でも毎回作り直した値と一致する()
        {
            var fonts = new[]
            {
                FontStyle.Default with { Name = "MS PGothic", SizePt = 10.5 },
                FontStyle.Default with { Name = "MS PGothic", SizePt = 10.5, Italic = true },
                FontStyle.Default with { Name = "MS PGothic", SizePt = 10.5, Color = new ArgbColor(255, 255, 0, 0) },
                FontStyle.Default with { SizePt = 7.25 },
            };

            foreach (var font in fonts)
            {
                var first = _metrics.GetMetrics(font);
                var second = _metrics.GetMetrics(font);

                using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
                var fresh = new SkiaFontMetricsProvider(resolver).GetMetrics(font);

                Assert.Equal(fresh, first);
                Assert.Equal(fresh, second);
            }
        }

        [Fact]
        public void フォントメトリクスの記録数に上限があっても正しい値を返す()
        {
            var font = FontStyle.Default with { Name = "MS PGothic" };
            var expected = new List<Utsushi.Layout.Text.FontMetrics>();
            var count = SkiaFontMetricsProvider.MaxCachedMetrics + 50;
            for (var i = 0; i < count; i++)
            {
                expected.Add(_metrics.GetMetrics(font with { SizePt = 1.0 + (i * 0.01) }));
            }

            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
            var fresh = new SkiaFontMetricsProvider(resolver);
            for (var i = 0; i < count; i += 97)
            {
                var size = 1.0 + (i * 0.01);
                Assert.Equal(fresh.GetMetrics(font with { SizePt = size }), _metrics.GetMetrics(font with { SizePt = size }));
                Assert.Equal(expected[i], _metrics.GetMetrics(font with { SizePt = size }));
            }
        }

        [Fact]
        public void 厳格モードで見つからないフォントは記録の有無によらず毎回エラーになる()
        {
            using var resolver = new FontResolver(FontResolverOptions.Strict);
            var metrics = new SkiaFontMetricsProvider(resolver);
            var font = FontStyle.Default with { Name = "Utsushi Missing Font For Test" };

            Assert.Throws<Utsushi.Core.Exceptions.FontNotAvailableException>(() => metrics.GetMetrics(font));
            Assert.Throws<Utsushi.Core.Exceptions.FontNotAvailableException>(() => metrics.GetMetrics(font));
        }

        [Fact]
        public void 描画用のフォントは同じ書体とサイズで使い回す()
        {
            using var context = RenderContext.Prepare(Layout(1, Array.Empty<DrawCommand>()), _metrics.Shaper, PdfRenderOptions.Default);
            var face = _fontResolver.Resolve(FontStyle.Default);

            var first = context.GetFont(face, 10.0);

            Assert.Same(first, context.GetFont(face, 10.0));
            Assert.NotSame(first, context.GetFont(face, 11.0));
            Assert.Equal(10f, first.Size);
        }

        [Fact]
        public void 画像は同じバイナリならデコードを1回で済ませる()
        {
            using var context = RenderContext.Prepare(Layout(1, Array.Empty<DrawCommand>()), _metrics.Shaper, PdfRenderOptions.Default);
            var png = SolidPng(8, 8);
            var decodes = 0;
            SKImage Decode(byte[] data)
            {
                decodes++;
                using var bitmap = SKBitmap.Decode(data);
                return SKImage.FromBitmap(bitmap);
            }

            var first = context.GetImage(png, Decode);
            var second = context.GetImage(png, Decode);
            context.GetImage((byte[])png.Clone(), Decode);

            Assert.Same(first, second);
            Assert.Equal(2, decodes);
        }

        [Fact]
        public void 全ページに同じ画像があってもPDFに埋め込む画像は1つで済む()
        {
            var png = NoisePng(200, 120);
            var commands = new DrawCommand[]
            {
                new ImageCommand(new RectPt(20, 20, 100, 60), png, "image/png", 0),
                new LineCommand(new PointPt(10, 100), new PointPt(200, 100), ArgbColor.Black, 0.75, LineDashStyle.Dash),
            };
            var renderer = new SkiaPdfRenderer(_metrics);

            var onePage = RenderSize(renderer, Layout(1, commands));
            var tenPages = RenderSize(renderer, Layout(10, commands));

            // 画像を1ページごとに埋め込むと、1ページ増えるごとに画像1枚分(数十KB)ずつ大きくなる。
            Assert.True(tenPages - onePage < png.Length, $"1ページ: {onePage} バイト, 10ページ: {tenPages} バイト, 画像: {png.Length} バイト");
        }

        private static long RenderSize(SkiaPdfRenderer renderer, PagedLayout layout)
        {
            using var stream = new MemoryStream();
            renderer.Render(layout, stream);
            return stream.Length;
        }

        private static PagedLayout Layout(int pageCount, IReadOnlyList<DrawCommand> commands)
        {
            var pages = new List<PageLayout>();
            for (var i = 1; i <= pageCount; i++)
            {
                pages.Add(new PageLayout(
                    PaperSize.A4, PageOrientation.Portrait, PaperSize.A4.WidthPt, PaperSize.A4.HeightPt,
                    commands, i, (1, 1), (1, 1), 1.0));
            }

            return new PagedLayout(pages, "test-report", "テストシート");
        }

        private static byte[] SolidPng(int width, int height)
        {
            using var bitmap = new SKBitmap(width, height);
            bitmap.Erase(SKColors.SteelBlue);
            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            return data.ToArray();
        }

        /// <summary>圧縮の効きにくい(PDFに埋め込むと大きくなる)画像。</summary>
        private static byte[] NoisePng(int width, int height)
        {
            var random = new Random(1);
            using var bitmap = new SKBitmap(width, height);
            for (var y = 0; y < height; y++)
            {
                for (var x = 0; x < width; x++)
                {
                    bitmap.SetPixel(x, y, new SKColor((byte)random.Next(256), (byte)random.Next(256), (byte)random.Next(256)));
                }
            }

            using var image = SKImage.FromBitmap(bitmap);
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            return data.ToArray();
        }
    }
}
