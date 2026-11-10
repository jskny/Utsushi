using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// フォント解決の検証(タスク6.2、および太字/斜体の合成判定)。
    /// </summary>
    public sealed class FontResolverTests
    {
        private static FontStyle Font(string name, bool bold = false, bool italic = false) =>
            FontStyle.Default with { Name = name, Bold = bold, Italic = italic };

        [Fact]
        public void 厳格モードでは未検出フォントを例外にする()
        {
            using var resolver = new FontResolver(FontResolverOptions.Strict);

            var ex = Assert.Throws<FontNotAvailableException>(
                () => resolver.Resolve(Font("存在しないフォント-ZZZ")));

            Assert.Equal("存在しないフォント-ZZZ", ex.FontName);
            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
        }

        [Fact]
        public void フォールバック許容なら代替フォントを返す()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            var resolved = resolver.Resolve(Font("存在しないフォント-ZZZ"));

            Assert.NotNull(resolved.Typeface);
        }

        [Fact]
        public void 代替したフォントは代替の書体名を返す()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            var family = resolver.GetSubstitution(Font("存在しないフォント-ZZZ", bold: true));

            Assert.Equal(Fonts.BundledFonts.JapaneseGothicFamily, family);
        }

        [Fact]
        public void 代替の書体名は解決済みのフォントでも返す()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());
            resolver.Resolve(Font("存在しないフォント-ZZZ"));

            Assert.Equal(Fonts.BundledFonts.JapaneseGothicFamily, resolver.GetSubstitution(Font("存在しないフォント-ZZZ")));
            Assert.Equal(Fonts.BundledFonts.JapaneseGothicFamily, resolver.GetSubstitution(Font("存在しないフォント-ZZZ")));
        }

        [Fact]
        public void 指定した名前で見つかったフォントは代替にならない()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            Assert.Null(resolver.GetSubstitution(Font(Fonts.BundledFonts.JapaneseGothicFamily)));
        }

        [Fact]
        public void 厳格モードでは代替せず問い合わせも例外になる()
        {
            using var resolver = new FontResolver(FontResolverOptions.Strict);

            Assert.Null(resolver.GetSubstitution(Font(Fonts.BundledFonts.JapaneseGothicFamily)));
            Assert.Throws<FontNotAvailableException>(() => resolver.GetSubstitution(Font("存在しないフォント-ZZZ")));
        }

        [Fact]
        public void 外字の補いは代替として返さない()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            // 外字用の代替書体(IPAmj明朝など)が入っていなくても、外字の補いの探索は代替として記録しない(要件3.4)。
            resolver.ResolveGlyphFallbacks(bold: false, italic: false);

            Assert.Null(resolver.GetSubstitution(Font(Fonts.BundledFonts.JapaneseGothicFamily)));
        }

        [Fact]
        public void 事前検証で未検出フォントをまとめて検出できる()
        {
            using var resolver = new FontResolver(FontResolverOptions.Strict);

            var ex = Assert.Throws<FontNotAvailableException>(
                () => resolver.EnsureAvailable(new[] { Font("存在しないフォント-ZZZ") }));

            Assert.Equal("存在しないフォント-ZZZ", ex.FontName);
        }

        [Fact]
        public void 存在しないフォントファイルの登録は例外になる()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            Assert.Throws<FontNotAvailableException>(
                () => resolver.RegisterFontFile("Any", Path.Combine(Path.GetTempPath(), "no-such-font.ttf")));
        }

        [Fact]
        public void 明示登録したフォントは名前で解決できる()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            using var resolver = new FontResolver(Options(("社内帳票フォント", fontPath)));

            var resolved = resolver.Resolve(Font("社内帳票フォント"));

            Assert.NotNull(resolved.Typeface);
            Assert.False(resolved.SynthesizeBold);
            Assert.False(resolved.SynthesizeItalic);
        }

        [Fact]
        public void 明示登録が通常字形だけなら太字と斜体は合成する()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            using var resolver = new FontResolver(Options(("社内帳票フォント", fontPath)));

            var bold = resolver.Resolve(Font("社内帳票フォント", bold: true));
            var italic = resolver.Resolve(Font("社内帳票フォント", italic: true));

            Assert.True(bold.SynthesizeBold);
            Assert.False(bold.SynthesizeItalic);
            Assert.True(italic.SynthesizeItalic);
            Assert.False(italic.SynthesizeBold);
        }

        [Fact]
        public void 太字のフォントファイルを明示登録すれば合成しない()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            using var resolver = new FontResolver(
                Options(("社内帳票フォント", fontPath), ("社内帳票フォント:bold", fontPath)));

            var resolved = resolver.Resolve(Font("社内帳票フォント", bold: true));

            // 太字キーで登録されたファイルが使われるため、描画側での合成は不要になる。
            Assert.False(resolved.SynthesizeBold);
        }

        [Fact]
        public void 太字の字形を持たないフォントは合成太字になる()
        {
            // 実行環境に「太字の字形を持たないフォント」が無い場合は検証できないため読み飛ばす。
            var family = FindFamilyWithoutRealBold();
            if (family is null)
            {
                return;
            }

            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            var resolved = resolver.Resolve(Font(family, bold: true));

            Assert.True(
                resolved.SynthesizeBold,
                $"'{family}' は太字の字形を持たないため、描画側で合成する必要がある");
        }

        [Fact]
        public void 太字の字形を持つフォントは実字形を使う()
        {
            var family = FindFamilyWithRealBold();
            if (family is null)
            {
                return;
            }

            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            var resolved = resolver.Resolve(Font(family, bold: true));

            Assert.False(
                resolved.SynthesizeBold,
                $"'{family}' は太字の字形を持つため、実字形をそのまま使うべき");
        }

        [Fact]
        public void 同じフォント指定は同じ解決結果を返す()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            var first = resolver.Resolve(Font("DejaVu Sans"));
            var second = resolver.Resolve(Font("DejaVu Sans"));

            Assert.Same(first.Typeface, second.Typeface);
        }

        [Fact]
        public void 不正なフォントキーは例外になる()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            Assert.Throws<ArgumentException>(() => resolver.RegisterFontFile("Family:heavy", fontPath));
        }

        private static FontResolverOptions Options(params (string Key, string Path)[] fontFiles) =>
            new(
                FontResolutionMode.Strict,
                fontFiles.ToDictionary(f => f.Key, f => f.Path, StringComparer.Ordinal),
                Array.Empty<string>(),
                null);

        [Fact]
        public void 指定した代替フォント名が見つからなければ作る時点でエラーにする()
        {
            // 同梱フォントで黙って代替せず、書体名の誤りに気付けるようにする(要件11.7)。
            var ex = Assert.Throws<FontNotAvailableException>(
                () => new FontResolver(FontResolverOptions.AllowFallback("存在しない書体-ZZZ")));

            Assert.Equal("存在しない書体-ZZZ", ex.FontName);
            Assert.Contains("存在しない書体-ZZZ", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 同梱フォントの名前は代替フォント名に指定できる()
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback("BIZ UDPGothic"));

            Assert.Equal("BIZ UDPGothic", resolver.GetSubstitution(FontStyle.Default with { Name = "存在しないフォント-ZZZ" }));
        }

        [Fact]
        public void 明示登録した書体は代替フォント名に指定できる()
        {
            var options = FontResolverOptions.AllowFallback("社内書体") with
            {
                FontFiles = new Dictionary<string, string> { ["社内書体"] = BundledFontFile() },
            };

            using var resolver = new FontResolver(options);

            // 代替の時点でも明示登録の書体を使う(同梱フォントの共有のインスタンスではなく、登録したファイルの書体)。
            var resolved = resolver.Resolve(FontStyle.Default with { Name = "存在しないフォント-ZZZ" });
            Assert.NotSame(Fonts.BundledFonts.JapaneseGothicTypeface, resolved.Typeface);
        }

        [Fact]
        public void 代替フォント名は日本語名と英語名の別名でも一致させる()
        {
            // 同梱フォントの日本語名(要件11.6)。
            using (new FontResolver(FontResolverOptions.AllowFallback("BIZ UDPゴシック")))
            {
            }

            // 英語名で明示登録した書体を、日本語名の代替フォント名で引く。
            var options = FontResolverOptions.AllowFallback("BIZ UDPゴシック") with
            {
                FontFiles = new Dictionary<string, string> { ["BIZ UDPGothic"] = BundledFontFile() },
            };
            using var resolver = new FontResolver(options);
            var resolved = resolver.Resolve(FontStyle.Default with { Name = "存在しないフォント-ZZZ" });
            Assert.NotSame(Fonts.BundledFonts.JapaneseGothicTypeface, resolved.Typeface);
        }

        [Fact]
        public void 厳格モードでは代替フォント名を確かめない()
        {
            using var resolver = new FontResolver(FontResolverOptions.Strict with { FallbackFamilyName = "存在しない書体-ZZZ" });
        }

        /// <summary>リポジトリにある同梱フォントのファイル(明示登録のテストに使う)。</summary>
        private static string BundledFontFile()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Utsushi.sln")))
            {
                directory = directory.Parent;
            }

            return Path.Combine(
                directory?.FullName ?? throw new InvalidOperationException("リポジトリのルートが見つかりません。"),
                "src", "Utsushi.Rendering", "Fonts", "BIZUDPGothic-Regular.ttf");
        }

        private static string? FindAnyInstalledFont() =>
            EnumerateFontFiles().FirstOrDefault();

        private static IEnumerable<string> EnumerateFontFiles()
        {
            foreach (var directory in new[] { "/usr/share/fonts", "/usr/local/share/fonts" })
            {
                if (!Directory.Exists(directory))
                {
                    continue;
                }

                foreach (var path in Directory.EnumerateFiles(directory, "*.ttf", SearchOption.AllDirectories))
                {
                    yield return path;
                }
            }
        }

        /// <summary>太字の字形を別ファイルとして持つファミリを探す。見つからなければ null。</summary>
        private static string? FindFamilyWithRealBold() => FindFamily(wantRealBold: true);

        /// <summary>太字の字形を持たないファミリを探す。見つからなければ null。</summary>
        private static string? FindFamilyWithoutRealBold() => FindFamily(wantRealBold: false);

        private static string? FindFamily(bool wantRealBold)
        {
            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            foreach (var family in new[] { "DejaVu Sans", "FreeSans", "IPAGothic", "IPAMincho", "DejaVu Serif" })
            {
                var regular = resolver.Resolve(Font(family));
                if (!string.Equals(regular.Typeface.FamilyName, family, StringComparison.OrdinalIgnoreCase))
                {
                    // このファミリは実行環境に存在しない。
                    continue;
                }

                var bold = resolver.Resolve(Font(family, bold: true));
                if (bold.SynthesizeBold != wantRealBold)
                {
                    return family;
                }
            }

            return null;
        }
    }
}
