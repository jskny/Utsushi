using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// フォントの登録と解決キャッシュの関係、およびコンストラクタが失敗した場合の後始末を確かめる。
    /// </summary>
    public sealed class FontResolverRegistrationTests
    {
        private static FontStyle Font(string name) => FontStyle.Default with { Name = name };

        [Fact]
        public void 解決の後に登録したフォントがそのキーに反映される()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            using var resolver = new FontResolver(FontResolverOptions.AllowFallback());

            // まだ登録していないため代替フォント(同梱のBIZ UDPゴシック)で解決され、キャッシュされる。
            var before = resolver.Resolve(Font("後から登録するフォント-XYZ"));

            resolver.RegisterFontFile("後から登録するフォント-XYZ", fontPath);
            var after = resolver.Resolve(Font("後から登録するフォント-XYZ"));

            Assert.NotSame(before.Typeface, after.Typeface);
            Assert.NotEqual(before.Typeface.FamilyName, after.Typeface.FamilyName);
        }

        [Fact]
        public void 厳格モードで見つからなかったフォントも登録後は解決できる()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            using var resolver = new FontResolver(FontResolverOptions.Strict);
            Assert.Throws<FontNotAvailableException>(() => resolver.Resolve(Font("後から登録するフォント-XYZ")));

            resolver.RegisterFontFile("後から登録するフォント-XYZ", fontPath);

            Assert.NotNull(resolver.Resolve(Font("後から登録するフォント-XYZ")).Typeface);
        }

        [Fact]
        public void 解決の後にディレクトリを登録すると外字用の代替書体の一覧も作り直される()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                File.Copy(fontPath, Path.Combine(directory, Path.GetFileName(fontPath)));

                using var resolver = new FontResolver(FontResolverOptions.Strict);
                var before = resolver.ResolveGlyphFallbacks(bold: false, italic: false);

                resolver.RegisterFontDirectory(directory);
                var after = resolver.ResolveGlyphFallbacks(bold: false, italic: false);

                // キャッシュを捨てていれば、同じ内容でも別のリストとして作り直される。
                Assert.NotSame(before, after);
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void コンストラクタが途中で失敗したら登録済みの書体を解放する()
        {
            var fontPath = FindAnyInstalledFont();
            if (fontPath is null)
            {
                return;
            }

            var options = new FontResolverOptions(
                FontResolutionMode.Strict,
                new Dictionary<string, string>
                {
                    ["登録できるフォント"] = fontPath,
                    ["存在しないフォント"] = Path.Combine(Path.GetTempPath(), "no-such-font-" + Guid.NewGuid().ToString("N") + ".ttf"),
                },
                Array.Empty<string>(),
                null);

            FontResolver? failed = null;
            FontResolver.ConstructionFailedForTesting = r => failed = r;
            try
            {
                Assert.Throws<FontNotAvailableException>(() => new FontResolver(options));
            }
            finally
            {
                FontResolver.ConstructionFailedForTesting = null;
            }

            Assert.NotNull(failed);
            Assert.Equal(0, failed!.OwnedTypefaceCountForTesting);
            Assert.Throws<ObjectDisposedException>(() => failed.Resolve(Font("登録できるフォント")));
        }

        private static string? FindAnyInstalledFont()
        {
            foreach (var directory in new[] { "/usr/share/fonts", "/usr/local/share/fonts" })
            {
                if (Directory.Exists(directory))
                {
                    var path = Directory.EnumerateFiles(directory, "*.ttf", SearchOption.AllDirectories).FirstOrDefault();
                    if (path is not null)
                    {
                        return path;
                    }
                }
            }

            return null;
        }
    }
}
