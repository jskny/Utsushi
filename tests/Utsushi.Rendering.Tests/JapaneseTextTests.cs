using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using SkiaSharp;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Utsushi.Rendering.Fonts;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// 日本語の人名・住所の文字(要件11): 同梱フォント、外字の文字単位の代替、異体字、濁点の合成、
    /// フォントのサブセット埋め込み、日本語のフォント名。
    /// </summary>
    /// <remarks>
    /// 実行環境のフォント構成に左右されないよう、同梱のBIZ UDPゴシックと、そこから作ったテスト用のフォントだけを使う
    /// (IPAmj明朝が必要なテストは、インストールされていない環境では何もせずに終える)。
    /// </remarks>
    public sealed class JapaneseTextTests : IDisposable
    {
        /// <summary>異体字セレクタ IVS の1番目(U+E0100)。BIZ UDPゴシックは「葛」+U+E0100 を通常と異なる字形で登録している。</summary>
        private const string KuzuWithIvs = "葛\U000E0100";

        private readonly string _tempDirectory = Path.Combine(Path.GetTempPath(), "utsushi-font-" + Guid.NewGuid().ToString("N"));

        public JapaneseTextTests() => Directory.CreateDirectory(_tempDirectory);

        public void Dispose()
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }

        private static FontStyle Font(string name, bool bold = false) =>
            FontStyle.Default with { Name = name, SizePt = 11.0, Bold = bold };

        private static FontResolverOptions BundledOnly(FontResolverOptions options) =>
            options with { GlyphFallbackFamilies = new[] { BundledFonts.JapaneseGothicFamily } };

        /// <summary>同梱フォントから英数字だけを取り出したテスト用フォント(日本語の字形を持たないセルのフォントの代わり)。</summary>
        private string CreateLatinOnlyFont()
        {
            var typeface = BundledFonts.JapaneseGothicTypeface;
            var map = new Dictionary<int, ushort>();
            for (var c = 0x20; c < 0x7F; c++)
            {
                map[c] = typeface.GetGlyph(c);
            }

            var subset = TrueTypeSubsetter.TryCreate(typeface, map.Values, map)!;
            var path = Path.Combine(_tempDirectory, "latin-only.ttf");
            File.WriteAllBytes(path, subset.FontData);
            return path;
        }

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

        private static TextCommand Text(string text, FontStyle font, double y = 50) =>
            new(new PointPt(50, y), text, font, TextAnchor.Left, null);

        private static byte[] Render(FontResolverOptions options, params DrawCommand[] commands)
        {
            using var resolver = new FontResolver(options);
            var renderer = new SkiaPdfRenderer(new SkiaFontMetricsProvider(resolver));
            using var output = new MemoryStream();
            renderer.Render(Layout(commands), output);
            return output.ToArray();
        }

        private static int CountEmbeddedFonts(byte[] pdf) =>
            Regex.Matches(System.Text.Encoding.Latin1.GetString(pdf), "/FontFile2").Count;

        // -- 要件11.1: 同梱の日本語フォント --------------------------------------

        [Fact]
        public void 代替フォント名を省略すると同梱の日本語フォントで代替する()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            var resolved = resolver.Resolve(Font("存在しないフォント-ZZZ"));

            Assert.NotEqual(0, resolved.Typeface.GetGlyph('請'));
        }

        [Theory]
        [InlineData("BIZ UDPGothic")]
        [InlineData("BIZ UDPゴシック")]
        public void 同梱フォントは厳格モードでも名前で解決できる(string name)
        {
            using var resolver = new FontResolver(FontResolverOptions.Strict);

            var resolved = resolver.Resolve(Font(name));

            Assert.NotEqual(0, resolved.Typeface.GetGlyph('髙'));
        }

        // -- 要件11.6: 日本語のフォント名 --------------------------------------

        [Theory]
        [InlineData("ＭＳ Ｐゴシック", "MS PGothic")]
        [InlineData("MS Pゴシック", "MS PGothic")]
        [InlineData("游ゴシック", "Yu Gothic")]
        [InlineData("Meiryo", "メイリオ")]
        public void 日本語のフォント名と英語のフォント名を同じフォントとみなす(string left, string right)
        {
            Assert.True(FontFamilyAliases.AreSame(left, right));
        }

        [Fact]
        public void 別のフォントは同じとみなさない()
        {
            Assert.False(FontFamilyAliases.AreSame("ＭＳ Ｐゴシック", "MS Gothic"));
        }

        // -- 要件11.2: 外字の文字単位の代替 --------------------------------------

        [Fact]
        public void セルのフォントに無い文字だけを外字用の代替フォントで描く()
        {
            var options = BundledOnly(FontResolverOptions.Strict with
            {
                FontFiles = new Dictionary<string, string> { ["LatinOnly"] = CreateLatinOnlyFont() },
            });
            using var resolver = new FontResolver(options);
            var shaper = new GlyphShaper(resolver);

            var shaped = shaper.Shape(Font("LatinOnly"), "No.髙橋");

            Assert.Null(shaped.Missing);
            Assert.Equal(2, shaped.Runs.Count);
            Assert.Same(shaped.Primary.Typeface, shaped.Runs[0].Face.Typeface);
            // 同名のフォントがインストールされていればそちらを、無ければ同梱フォントを使う。
            Assert.Equal(BundledFonts.JapaneseGothicFamily, shaped.Runs[1].Face.Typeface.FamilyName);
        }

        [Fact]
        public void 外字用の代替フォントで描く文字の送り幅も計測に含める()
        {
            var options = BundledOnly(FontResolverOptions.Strict with
            {
                FontFiles = new Dictionary<string, string> { ["LatinOnly"] = CreateLatinOnlyFont() },
            });
            using var resolver = new FontResolver(options);
            var metrics = new SkiaFontMetricsProvider(resolver);

            var latin = metrics.MeasureTextWidth(Font("LatinOnly"), "No.");
            var mixed = metrics.MeasureTextWidth(Font("LatinOnly"), "No.髙橋");

            // 全角2文字ぶん(11pt前後 × 2)広がる。
            Assert.InRange(mixed - latin, 18.0, 24.0);
        }

        [Fact]
        public void 太字の指定は外字用の代替フォントにも引き継ぐ()
        {
            var options = BundledOnly(FontResolverOptions.Strict with
            {
                FontFiles = new Dictionary<string, string> { ["LatinOnly"] = CreateLatinOnlyFont() },
            });
            using var resolver = new FontResolver(options);

            var shaped = new GlyphShaper(resolver).Shape(Font("LatinOnly", bold: true), "A髙");

            // 実字形の太字がインストールされていればそれを、無ければ描画側で合成する。
            var face = shaped.Runs[1].Face;
            Assert.True(face.SynthesizeBold || face.Typeface.FontWeight >= (int)SKFontStyleWeight.SemiBold);
        }

        [Fact]
        public void 外字用の代替フォントを空にすると文字単位の代替を行わない()
        {
            var options = FontResolverOptions.Strict with
            {
                FontFiles = new Dictionary<string, string> { ["LatinOnly"] = CreateLatinOnlyFont() },
                GlyphFallbackFamilies = Array.Empty<string>(),
            };

            var ex = Assert.Throws<MissingGlyphException>(() => Render(options, Text("No.髙橋", Font("LatinOnly"))));

            Assert.Equal('髙', ex.CodePoint);
        }

        [Fact]
        public void どのフォントにも無い外字はエラーになる()
        {
            // 「𠮷」(U+20BB7)はJIS第1〜第4水準の外にあり、同梱のBIZ UDPゴシックは持たない。
            var ex = Assert.Throws<MissingGlyphException>(
                () => Render(BundledOnly(FontResolverOptions.AllowFallback()), Text("𠮷野 様", Font("MS PGothic"))));

            Assert.Equal(0x20BB7, ex.CodePoint);
        }

        [Fact]
        public void IPAmj明朝がインストールされていれば既定で外字に使う()
        {
            using var probe = SKTypeface.FromFamilyName("IPAmjMincho");
            if (probe is null || probe.FamilyName != "IPAmjMincho")
            {
                return; // IPAmj明朝の無い環境(CIなど)では確認できない
            }

            var pdf = Render(FontResolverOptions.AllowFallback(), Text("𠮷野 様", Font("MS PGothic")));

            Assert.Equal(2, CountEmbeddedFonts(pdf));
            Assert.True(pdf.Length < 200_000, $"IPAmj明朝(46MB)もサブセット化されるはず (PDF: {pdf.Length} bytes)");
        }

        // -- 要件11.3: 異体字 --------------------------------------------------

        [Fact]
        public void 異体字セレクタ付きの文字はフォントに登録された異体字の字形で描く()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
            var shaper = new GlyphShaper(resolver);

            var normal = shaper.Shape(Font(BundledFonts.JapaneseGothicFamily), "葛");
            var variant = shaper.Shape(Font(BundledFonts.JapaneseGothicFamily), KuzuWithIvs);

            Assert.Null(variant.Missing);
            var glyph = Assert.Single(Assert.Single(variant.Runs).Glyphs);
            Assert.NotEqual(0, glyph);
            Assert.NotEqual(normal.Runs[0].Glyphs[0], glyph);
        }

        [Fact]
        public void どのフォントにも登録されていない異体字はエラーになる()
        {
            var ex = Assert.Throws<MissingGlyphException>(
                () => Render(BundledOnly(FontResolverOptions.AllowFallback()), Text("葛\U000E01EF", Font("MS PGothic"))));

            Assert.Equal(0xE01EF, ex.CodePoint);
            Assert.Contains("異体字", ex.Message);
        }

        [Fact]
        public void 描画を続ける設定なら未登録の異体字は基底文字で描く()
        {
            using var resolver = new FontResolver(BundledOnly(FontResolverOptions.AllowFallback()));

            var shaped = new GlyphShaper(resolver).Shape(Font("MS PGothic"), "葛\U000E01EF");

            Assert.NotNull(shaped.Missing);
            Assert.Equal(resolver.Resolve(Font("MS PGothic")).Typeface.GetGlyph('葛'), Assert.Single(shaped.Runs[0].Glyphs));
        }

        [Fact]
        public void 同じ文書で通常の字形と異体字を使うと別々のサブセットに分けて埋め込む()
        {
            // cmapは1文字に1字形しか対応づけられないため、どちらも「葛」として文字列抽出できるよう分ける。
            var font = Font(BundledFonts.JapaneseGothicFamily);
            var pdf = Render(FontResolverOptions.AllowFallback(), Text("葛西", font), Text(KuzuWithIvs + "飾区", font, 80));

            Assert.Equal(2, CountEmbeddedFonts(pdf));
        }

        // -- 要件11.4: 濁点・半濁点の合成、CJK互換漢字 ---------------------------

        [Theory]
        [InlineData("が", 'が')]
        [InlineData("パ", 'パ')]
        public void 仮名と合成用の濁点は合成済みの仮名で描く(string decomposed, char composed)
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
            var font = Font(BundledFonts.JapaneseGothicFamily);

            var shaped = new GlyphShaper(resolver).Shape(font, decomposed);

            Assert.Equal(resolver.Resolve(font).Typeface.GetGlyph(composed), Assert.Single(Assert.Single(shaped.Runs).Glyphs));
        }

        [Fact]
        public void CJK互換漢字は統合漢字に置き換えない()
        {
            // 「﨑」(U+FA11)は「崎」(U+5D0E)と別の字形。正規化すると人名の字形が変わってしまう。
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
            var font = Font(BundledFonts.JapaneseGothicFamily);
            var typeface = resolver.Resolve(font).Typeface;

            var shaped = new GlyphShaper(resolver).Shape(font, "﨑");

            Assert.Equal(typeface.GetGlyph(0xFA11), Assert.Single(shaped.Runs[0].Glyphs));
            Assert.NotEqual(typeface.GetGlyph(0x5D0E), shaped.Runs[0].Glyphs[0]);
        }

        // -- 要件11.5: サブセット埋め込み --------------------------------------

        [Fact]
        public void 埋め込むフォントは使った字形だけのサブセットにする()
        {
            var pdf = Render(FontResolverOptions.AllowFallback(), Text("請求書 株式会社サンプル 御中", Font("MS PGothic")));

            // 同梱フォント全体(約4.6MB)を埋め込むと数MBになる。
            Assert.True(pdf.Length < 100_000, $"PDF: {pdf.Length} bytes");
            Assert.Equal(1, CountEmbeddedFonts(pdf));
        }

        [Fact]
        public void サブセットの字形と送り幅は元のフォントと同じ()
        {
            var typeface = BundledFonts.JapaneseGothicTypeface;
            const string text = "髙橋 﨑 No.1";
            var map = text.EnumerateRunes().Distinct().ToDictionary(r => r.Value, r => typeface.GetGlyph(r.Value));

            var subset = TrueTypeSubsetter.TryCreate(typeface, map.Values, map)!;
            using var data = SKData.CreateCopy(subset.FontData);
            using var subsetTypeface = SKTypeface.FromData(data);
            using var original = new SKFont(typeface, 20);
            using var reduced = new SKFont(subsetTypeface, 20);

            Assert.True(subset.FontData.Length < 50_000, $"サブセット: {subset.FontData.Length} bytes");
            foreach (var (codePoint, glyph) in map)
            {
                var mapped = subset.MapGlyph(glyph);
                Assert.Equal(mapped, subsetTypeface.GetGlyph(codePoint));
                Assert.Equal(original.MeasureText(new[] { glyph }), reduced.MeasureText(new[] { mapped }), 3);
                using var originalPath = original.GetGlyphPath(glyph);
                using var subsetPath = reduced.GetGlyphPath(mapped);
                Assert.Equal(originalPath.Bounds, subsetPath.Bounds);
            }
        }

        [Fact]
        public void アウトライン出力ではフォントを埋め込まない()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
            var renderer = new SkiaPdfRenderer(new SkiaFontMetricsProvider(resolver), PdfRenderOptions.OutlineText);
            using var output = new MemoryStream();

            renderer.Render(Layout(Text("髙橋 " + KuzuWithIvs, Font("MS PGothic"))), output);

            Assert.Equal(0, CountEmbeddedFonts(output.ToArray()));
        }
    }
}
