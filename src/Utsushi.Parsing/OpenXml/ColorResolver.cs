using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Utsushi.Core;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// OOXML の色指定(rgb / indexed / theme+tint)を <see cref="ArgbColor"/> へ解決する。
    /// </summary>
    internal sealed class ColorResolver
    {
        /// <summary>
        /// Excel の既定インデックスカラーパレット。
        /// ブックに <c>indexedColors</c> が無い場合に使う(ECMA-376 の既定値)。
        /// </summary>
        private static readonly uint[] DefaultIndexedPalette =
        {
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF000000, 0xFFFFFFFF, 0xFFFF0000, 0xFF00FF00, 0xFF0000FF, 0xFFFFFF00, 0xFFFF00FF, 0xFF00FFFF,
        0xFF800000, 0xFF008000, 0xFF000080, 0xFF808000, 0xFF800080, 0xFF008080, 0xFFC0C0C0, 0xFF808080,
        0xFF9999FF, 0xFF993366, 0xFFFFFFCC, 0xFFCCFFFF, 0xFF660066, 0xFFFF8080, 0xFF0066CC, 0xFFCCCCFF,
        0xFF000080, 0xFFFF00FF, 0xFFFFFF00, 0xFF00FFFF, 0xFF800080, 0xFF800000, 0xFF008080, 0xFF0000FF,
        0xFF00CCFF, 0xFFCCFFFF, 0xFFCCFFCC, 0xFFFFFF99, 0xFF99CCFF, 0xFFFF99CC, 0xFFCC99FF, 0xFFFFCC99,
        0xFF3366FF, 0xFF33CCCC, 0xFF99CC00, 0xFFFFCC00, 0xFFFF9900, 0xFFFF6600, 0xFF666699, 0xFF969696,
        0xFF003366, 0xFF339966, 0xFF003300, 0xFF333300, 0xFF993300, 0xFF993366, 0xFF333399, 0xFF333333,
    };

        /// <summary>インデックス64/65は「自動」(前景=黒 / 背景=白)を意味する。</summary>
        private const int AutomaticForegroundIndex = 64;
        private const int AutomaticBackgroundIndex = 65;

        private readonly IReadOnlyList<uint> _indexedPalette;
        private readonly IReadOnlyList<ArgbColor> _themeColors;

        public ColorResolver(IReadOnlyList<uint> indexedPalette, IReadOnlyList<ArgbColor> themeColors)
        {
            _indexedPalette = indexedPalette;
            _themeColors = themeColors;
        }

        /// <summary>ワークブックパートからテーマ色・インデックスパレットを読み出してリゾルバを構築する。</summary>
        public static ColorResolver Create(WorkbookPart workbookPart)
        {
            var palette = ReadIndexedPalette(workbookPart);
            var theme = ReadThemeColors(workbookPart);
            return new ColorResolver(palette, theme);
        }

        /// <summary>色指定を解決する。解決できない場合は <paramref name="fallback"/> を返す。</summary>
        public ArgbColor Resolve(X.ColorType? color, ArgbColor fallback)
        {
            if (color is null)
            {
                return fallback;
            }

            if (color.Auto?.Value == true)
            {
                return fallback;
            }

            ArgbColor? baseColor = null;

            if (color.Rgb?.Value is { Length: > 0 } rgb && ArgbColor.TryParseHex(rgb, out var parsed))
            {
                baseColor = parsed;
            }
            else if (color.Indexed?.Value is { } indexed)
            {
                baseColor = ResolveIndexed((int)indexed, fallback);
            }
            else if (color.Theme?.Value is { } themeIndex)
            {
                baseColor = ResolveTheme((int)themeIndex, fallback);
            }

            if (baseColor is null)
            {
                return fallback;
            }

            var tint = color.Tint?.Value ?? 0.0;
            return Math.Abs(tint) < 1e-9 ? baseColor.Value : ApplyTint(baseColor.Value, tint);
        }

        private ArgbColor ResolveIndexed(int index, ArgbColor fallback)
        {
            if (index == AutomaticForegroundIndex)
            {
                return ArgbColor.Black;
            }

            if (index == AutomaticBackgroundIndex)
            {
                return ArgbColor.White;
            }

            if (index < 0 || index >= _indexedPalette.Count)
            {
                return fallback;
            }

            return FromUInt(_indexedPalette[index]);
        }

        private ArgbColor ResolveTheme(int index, ArgbColor fallback) =>
            index >= 0 && index < _themeColors.Count ? _themeColors[index] : fallback;

        /// <summary>
        /// OOXML の tint をHLS空間で適用する(ECMA-376 Part 1, 18.3.1.15 の定義に従う)。
        /// </summary>
        internal static ArgbColor ApplyTint(ArgbColor color, double tint)
        {
            var (h, l, s) = RgbToHls(color.R, color.G, color.B);

            // ECMA-376: tint<0 は Lum*(1+tint)、tint>0 は Lum*(1-tint)+tint(輝度の最大値を1.0に正規化した形)。
            l = tint < 0 ? l * (1.0 + tint) : (l * (1.0 - tint)) + tint;

            l = Math.Min(1.0, Math.Max(0.0, l));
            var (r, g, b) = HlsToRgb(h, l, s);
            return new ArgbColor(color.A, r, g, b);
        }

        private static (double H, double L, double S) RgbToHls(byte r8, byte g8, byte b8)
        {
            var r = r8 / 255.0;
            var g = g8 / 255.0;
            var b = b8 / 255.0;

            var max = Math.Max(r, Math.Max(g, b));
            var min = Math.Min(r, Math.Min(g, b));
            var l = (max + min) / 2.0;

            if (Math.Abs(max - min) < 1e-9)
            {
                return (0.0, l, 0.0);
            }

            var d = max - min;
            var s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

            double h;
            if (Math.Abs(max - r) < 1e-9)
            {
                h = ((g - b) / d) + (g < b ? 6.0 : 0.0);
            }
            else if (Math.Abs(max - g) < 1e-9)
            {
                h = ((b - r) / d) + 2.0;
            }
            else
            {
                h = ((r - g) / d) + 4.0;
            }

            return (h / 6.0, l, s);
        }

        private static (byte R, byte G, byte B) HlsToRgb(double h, double l, double s)
        {
            if (s <= 1e-9)
            {
                var v = ToByte(l);
                return (v, v, v);
            }

            var q = l < 0.5 ? l * (1.0 + s) : l + s - (l * s);
            var p = (2.0 * l) - q;
            return (ToByte(HueToChannel(p, q, h + (1.0 / 3.0))), ToByte(HueToChannel(p, q, h)), ToByte(HueToChannel(p, q, h - (1.0 / 3.0))));
        }

        private static double HueToChannel(double p, double q, double t)
        {
            if (t < 0) { t += 1.0; }
            if (t > 1) { t -= 1.0; }
            if (t < 1.0 / 6.0) { return p + ((q - p) * 6.0 * t); }
            if (t < 1.0 / 2.0) { return q; }
            if (t < 2.0 / 3.0) { return p + ((q - p) * ((2.0 / 3.0) - t) * 6.0); }
            return p;
        }

        private static byte ToByte(double value) => (byte)Math.Round(Math.Min(1.0, Math.Max(0.0, value)) * 255.0, MidpointRounding.AwayFromZero);

        private static ArgbColor FromUInt(uint value) =>
            new((byte)((value >> 24) & 0xFF), (byte)((value >> 16) & 0xFF), (byte)((value >> 8) & 0xFF), (byte)(value & 0xFF));

        private static IReadOnlyList<uint> ReadIndexedPalette(WorkbookPart workbookPart)
        {
            var colors = workbookPart.WorkbookStylesPart?.Stylesheet?.Colors?.IndexedColors;
            if (colors is null)
            {
                return DefaultIndexedPalette;
            }

            var result = new List<uint>();
            foreach (var rgbColor in colors.Elements<X.RgbColor>())
            {
                if (rgbColor.Rgb?.Value is { } hex
                    && uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value))
                {
                    result.Add(value);
                }
                else
                {
                    result.Add(0xFF000000);
                }
            }

            return result.Count > 0 ? result : DefaultIndexedPalette;
        }

        private static IReadOnlyList<ArgbColor> ReadThemeColors(WorkbookPart workbookPart)
        {
            var scheme = workbookPart.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            if (scheme is null)
            {
                return Array.Empty<ArgbColor>();
            }

            // Excel の theme インデックスは clrScheme の並び順と異なり、lt1/dk1・lt2/dk2 が入れ替わる。
            var ordered = new Color2Type?[]
            {
            scheme.Light1Color, scheme.Dark1Color, scheme.Light2Color, scheme.Dark2Color,
            scheme.Accent1Color, scheme.Accent2Color, scheme.Accent3Color,
            scheme.Accent4Color, scheme.Accent5Color, scheme.Accent6Color,
            scheme.Hyperlink, scheme.FollowedHyperlinkColor,
            };

            var result = new List<ArgbColor>(ordered.Length);
            foreach (var entry in ordered)
            {
                result.Add(ReadSchemeColor(entry));
            }

            return result;
        }

        private static ArgbColor ReadSchemeColor(Color2Type? color)
        {
            if (color is null)
            {
                return ArgbColor.Black;
            }

            if (color.RgbColorModelHex?.Val?.Value is { } hex && ArgbColor.TryParseHex(hex, out var rgb))
            {
                return rgb;
            }

            // sysClr(windowText/window)は既定のシステム色として扱う。
            if (color.SystemColor is { } sys)
            {
                if (sys.LastColor?.Value is { } last && ArgbColor.TryParseHex(last, out var lastColor))
                {
                    return lastColor;
                }

                return sys.Val is not null && sys.Val.Value == SystemColorValues.Window ? ArgbColor.White : ArgbColor.Black;
            }

            return ArgbColor.Black;
        }
    }
}
