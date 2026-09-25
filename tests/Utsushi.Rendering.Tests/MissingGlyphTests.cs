using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// 字形の無い文字の検出(要件5.5)と、フォント解決の並行実行(同じコンバータを複数スレッドから使う場合)。
    /// </summary>
    public sealed class MissingGlyphTests : IDisposable
    {
        /// <summary>
        /// Unicodeで未割り当てのコードポイント(ギリシア文字ブロック内の空き)。どのフォントにも字形が無い。
        /// </summary>
        private const string UnassignedCharacter = "͸";

        private readonly FontResolver _fontResolver;
        private readonly SkiaFontMetricsProvider _metrics;

        public MissingGlyphTests()
        {
            _fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            _metrics = new SkiaFontMetricsProvider(_fontResolver);
        }

        public void Dispose() => _fontResolver.Dispose();

        private static PagedLayout Layout(params DrawCommand[] commands) =>
            new(
                new[]
                {
                    new PageLayout(
                        PaperSize.A4, PageOrientation.Portrait, PaperSize.A4.WidthPt, PaperSize.A4.HeightPt,
                        commands, 1, (1, 1), (1, 1), 1.0),
                },
                "test-report",
                "テストシート");

        private static TextCommand Text(string text) =>
            new(new PointPt(50, 50), text, FontStyle.Default, TextAnchor.Left, null);

        [Theory]
        [InlineData(PdfTextRendering.EmbedFont)]
        [InlineData(PdfTextRendering.Outline)]
        public void 字形の無い文字があればPDFを出力せずエラーになる(PdfTextRendering textRendering)
        {
            var renderer = new SkiaPdfRenderer(_metrics, PdfRenderOptions.Default with { TextRendering = textRendering });
            using var output = new MemoryStream();

            var ex = Assert.Throws<MissingGlyphException>(
                () => renderer.Render(Layout(Text("ABC" + UnassignedCharacter)), output));

            Assert.Equal(0x0378, ex.CodePoint);
            Assert.Equal("ABC" + UnassignedCharacter, ex.Text);
            Assert.False(string.IsNullOrEmpty(ex.FontName));
            Assert.Equal("test-report", ex.ReportCode);
            Assert.Equal("テストシート", ex.SheetName);
            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Contains("U+0378", ex.Message);
            Assert.Equal(0, output.Length);
        }

        [Fact]
        public void 図形内テキストの字形の欠落も検出する()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var shape = new ShapeCommand(
                new RectPt(10, 10, 100, 40),
                ShapePresetType.Rect,
                Array.Empty<double>(),
                RotationDegrees: 0,
                Fill: null,
                Outline: new ShapeOutline(ArgbColor.Black, 1.0),
                TextLines: new[]
                {
                    new ShapeTextLine(new PointPt(20, 30), UnassignedCharacter, FontStyle.Default, TextAnchor.Left),
                });

            var ex = Assert.Throws<MissingGlyphException>(() => renderer.Render(Layout(shape), output));

            Assert.Equal("test-report", ex.ReportCode);
        }

        [Fact]
        public void 字形がすべて揃っていればエラーにならない()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            renderer.Render(Layout(Text("Invoice No. 0001")), output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void 描画を続ける設定なら字形が無くてもPDFを出力する()
        {
            var renderer = new SkiaPdfRenderer(
                _metrics, PdfRenderOptions.Default with { MissingGlyphs = MissingGlyphPolicy.Render });
            using var output = new MemoryStream();

            renderer.Render(Layout(Text("ABC" + UnassignedCharacter)), output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void 同じフォントを複数スレッドから同時に解決しても同じ書体が返る()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
            var font = FontStyle.Default with { Name = "並行解決テスト用の存在しないフォント" };

            var results = new ResolvedTypeface[64];
            Parallel.For(0, results.Length, i => results[i] = resolver.Resolve(font));

            Assert.All(results, r => Assert.Same(results[0], r));
        }

        [Fact]
        public void 同じレンダラーを複数スレッドから同時に使える()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            var layout = Layout(Text("Invoice No. 0001"));

            var sizes = new long[16];
            Parallel.For(0, sizes.Length, i =>
            {
                using var output = new MemoryStream();
                renderer.Render(layout, output);
                sizes[i] = output.Length;
            });

            Assert.All(sizes, size => Assert.True(size > 0));
        }
    }
}
