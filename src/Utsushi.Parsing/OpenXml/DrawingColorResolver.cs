using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Utsushi.Core;
using Dr = DocumentFormat.OpenXml.Drawing;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// 図形・接続線の色(DrawingML の <c>a:srgbClr</c>/<c>a:schemeClr</c>/<c>a:sysClr</c> と色の修飾)を解決する
    /// (要件10.15)。図形のスタイル参照(<c>lnRef</c>)の線の太さを引くため、テーマの <c>lnStyleLst</c> も保持する(要件10.16)。
    /// </summary>
    /// <remarks>
    /// セルの色(SpreadsheetML の theme/tint)は <see cref="ColorResolver"/> が扱う。DrawingML の <c>tint</c>/<c>shade</c> は
    /// SpreadsheetML の <c>tint</c> と計算方法が異なる(線形RGBで白/黒へ補間する)ため、別のクラスにしている。
    /// </remarks>
    internal sealed class DrawingColorResolver
    {
        /// <summary>1つの色に適用する修飾(<c>lumMod</c> 等)の個数の上限。信頼できない入力への安全弁。</summary>
        private const int MaxColorModifiers = 16;

        /// <summary>テーマの <c>lnStyleLst</c> から読む線の書式の個数の上限(Excelのテーマは3個)。</summary>
        private const int MaxLineStyles = 8;

        /// <summary>テーマの <c>lnStyleLst</c> に該当が無い場合の線の太さ(ポイント)。</summary>
        internal const double DefaultLineStyleWidthPt = 0.75;

        /// <summary>テーマが無い場合の <c>lnStyleLst</c> の線の太さ(Office 既定テーマ: 6350/12700/19050 EMU)。</summary>
        private static readonly IReadOnlyList<double> DefaultLineStyleWidthsPt = new[] { 0.5, 1.0, 1.5 };

        /// <summary>
        /// テーマが無い場合の配色(Office 2013〜の既定テーマ「Office」)。
        /// </summary>
        private static readonly IReadOnlyDictionary<string, ArgbColor> DefaultScheme = new Dictionary<string, ArgbColor>(StringComparer.Ordinal)
        {
            ["dk1"] = Rgb(0x000000),
            ["lt1"] = Rgb(0xFFFFFF),
            ["dk2"] = Rgb(0x44546A),
            ["lt2"] = Rgb(0xE7E6E6),
            ["accent1"] = Rgb(0x4472C4),
            ["accent2"] = Rgb(0xED7D31),
            ["accent3"] = Rgb(0xA5A5A5),
            ["accent4"] = Rgb(0xFFC000),
            ["accent5"] = Rgb(0x5B9BD5),
            ["accent6"] = Rgb(0x70AD47),
            ["hlink"] = Rgb(0x0563C1),
            ["folHlink"] = Rgb(0x954F72),
        };

        private readonly IReadOnlyDictionary<string, ArgbColor> _scheme;
        private readonly IReadOnlyList<double> _lineStyleWidthsPt;

        private DrawingColorResolver(IReadOnlyDictionary<string, ArgbColor> scheme, IReadOnlyList<double> lineStyleWidthsPt)
        {
            _scheme = scheme;
            _lineStyleWidthsPt = lineStyleWidthsPt;
        }

        /// <summary>テーマの無いブック用(既定の配色)。</summary>
        public static DrawingColorResolver Default { get; } = new(DefaultScheme, DefaultLineStyleWidthsPt);

        /// <summary>ブックのテーマ(<c>theme1.xml</c>)から配色と線の書式を読み込む。</summary>
        public static DrawingColorResolver Create(WorkbookPart workbookPart)
        {
            var theme = workbookPart.ThemePart?.Theme;
            var colorScheme = theme?.ThemeElements?.ColorScheme;
            if (colorScheme is null)
            {
                return Default;
            }

            var scheme = new Dictionary<string, ArgbColor>(StringComparer.Ordinal)
            {
                ["dk1"] = ColorResolver.ReadSchemeColor(colorScheme.Dark1Color),
                ["lt1"] = ColorResolver.ReadSchemeColor(colorScheme.Light1Color),
                ["dk2"] = ColorResolver.ReadSchemeColor(colorScheme.Dark2Color),
                ["lt2"] = ColorResolver.ReadSchemeColor(colorScheme.Light2Color),
                ["accent1"] = ColorResolver.ReadSchemeColor(colorScheme.Accent1Color),
                ["accent2"] = ColorResolver.ReadSchemeColor(colorScheme.Accent2Color),
                ["accent3"] = ColorResolver.ReadSchemeColor(colorScheme.Accent3Color),
                ["accent4"] = ColorResolver.ReadSchemeColor(colorScheme.Accent4Color),
                ["accent5"] = ColorResolver.ReadSchemeColor(colorScheme.Accent5Color),
                ["accent6"] = ColorResolver.ReadSchemeColor(colorScheme.Accent6Color),
                ["hlink"] = ColorResolver.ReadSchemeColor(colorScheme.Hyperlink),
                ["folHlink"] = ColorResolver.ReadSchemeColor(colorScheme.FollowedHyperlinkColor),
            };

            var widths = new List<double>();
            var lineStyles = theme!.ThemeElements!.FormatScheme?.LineStyleList;
            if (lineStyles is not null)
            {
                foreach (var outline in lineStyles.Elements<Dr.Outline>().Take(MaxLineStyles))
                {
                    widths.Add(outline.Width?.Value is { } w ? Units.EmusToPoints(w) : DefaultLineStyleWidthPt);
                }
            }

            return new DrawingColorResolver(scheme, widths);
        }

        /// <summary>
        /// <c>lnRef/@idx</c>(1始まり)が指すテーマの線の書式の太さ(ポイント)。該当が無ければ既定値。
        /// </summary>
        public double GetLineStyleWidthPt(uint index) =>
            index >= 1 && index <= _lineStyleWidthsPt.Count ? _lineStyleWidthsPt[(int)index - 1] : DefaultLineStyleWidthPt;

        /// <summary>
        /// 色の要素(<c>a:srgbClr</c>/<c>a:schemeClr</c>/<c>a:sysClr</c>)を子に持つ要素(<c>a:solidFill</c>・<c>a:gs</c>・
        /// <c>a:lnRef</c> 等)から色を求める。
        /// </summary>
        /// <param name="container">色の要素を子に持つ要素。</param>
        /// <param name="placeholder">
        /// <c>a:schemeClr val="phClr"</c> に差し込む色(テーマの書式設定の中で、スタイル参照の色を表す)。無ければ null。
        /// </param>
        /// <param name="color">求めた色。</param>
        /// <returns>色を求められた場合は true。<c>a:prstClr</c> など未対応の指定や、色の要素が無い場合は false。</returns>
        public bool TryResolve(OpenXmlElement? container, ArgbColor? placeholder, out ArgbColor color)
        {
            color = default;
            if (container is null)
            {
                return false;
            }

            foreach (var child in container.ChildElements)
            {
                switch (child)
                {
                    case Dr.RgbColorModelHex rgb:
                        if (!ArgbColor.TryParseHex(rgb.Val?.Value, out var baseRgb))
                        {
                            return false;
                        }

                        color = ApplyModifiers(baseRgb, rgb);
                        return true;

                    case Dr.SchemeColor scheme:
                        if (!TryResolveSchemeName(scheme.Val?.InnerText, placeholder, out var baseScheme))
                        {
                            return false;
                        }

                        color = ApplyModifiers(baseScheme, scheme);
                        return true;

                    case Dr.SystemColor system:
                        var baseSystem = ArgbColor.TryParseHex(system.LastColor?.Value, out var last)
                            ? last
                            : system.Val?.InnerText == "window" ? ArgbColor.White : ArgbColor.Black;
                        color = ApplyModifiers(baseSystem, system);
                        return true;
                }
            }

            return false;
        }

        private bool TryResolveSchemeName(string? name, ArgbColor? placeholder, out ArgbColor color)
        {
            color = default;
            switch (name)
            {
                case null:
                    return false;
                case "phClr":
                    if (placeholder is { } ph)
                    {
                        color = ph;
                        return true;
                    }

                    return false;
            }

            // Excel の既定の配色の対応(clrMap): bg1=lt1, tx1=dk1, bg2=lt2, tx2=dk2。
            var key = name switch
            {
                "bg1" => "lt1",
                "tx1" => "dk1",
                "bg2" => "lt2",
                "tx2" => "dk2",
                _ => name,
            };

            return _scheme.TryGetValue(key, out color);
        }

        /// <summary>色の要素の子にある修飾を、文書の順に適用する。</summary>
        private static ArgbColor ApplyModifiers(ArgbColor color, OpenXmlElement colorElement)
        {
            var r = color.R / 255.0;
            var g = color.G / 255.0;
            var b = color.B / 255.0;
            var a = color.A / 255.0;

            var count = 0;
            foreach (var modifier in colorElement.ChildElements)
            {
                if (++count > MaxColorModifiers)
                {
                    break;
                }

                switch (modifier)
                {
                    case Dr.LuminanceModulation lumMod when lumMod.Val?.Value is { } v:
                        ModifyLuminance(ref r, ref g, ref b, l => l * (v / 100000.0));
                        break;
                    case Dr.LuminanceOffset lumOff when lumOff.Val?.Value is { } v:
                        ModifyLuminance(ref r, ref g, ref b, l => l + (v / 100000.0));
                        break;
                    case Dr.Shade shade when shade.Val?.Value is { } v:
                        // 黒へ補間する(線形RGBで各成分に値を掛ける)。
                        r = ScaleLinear(r, c => c * Fraction(v));
                        g = ScaleLinear(g, c => c * Fraction(v));
                        b = ScaleLinear(b, c => c * Fraction(v));
                        break;
                    case Dr.Tint tint when tint.Val?.Value is { } v:
                        // 白へ補間する(線形RGBで 1 - (1 - c) × 値)。
                        r = ScaleLinear(r, c => 1.0 - ((1.0 - c) * Fraction(v)));
                        g = ScaleLinear(g, c => 1.0 - ((1.0 - c) * Fraction(v)));
                        b = ScaleLinear(b, c => 1.0 - ((1.0 - c) * Fraction(v)));
                        break;
                    case Dr.Alpha alpha when alpha.Val?.Value is { } v:
                        a = Fraction(v);
                        break;
                }
            }

            return new ArgbColor(ToByte(a), ToByte(r), ToByte(g), ToByte(b));
        }

        private static double Fraction(int value) => Clamp01(value / 100000.0);

        private static double Clamp01(double value) => double.IsNaN(value) ? 0.0 : Math.Max(0.0, Math.Min(1.0, value));

        private static byte ToByte(double value) => (byte)Math.Round(Clamp01(value) * 255.0);

        private static ArgbColor Rgb(int rgb) =>
            new(0xFF, (byte)((rgb >> 16) & 0xFF), (byte)((rgb >> 8) & 0xFF), (byte)(rgb & 0xFF));

        /// <summary>sRGB の成分を線形RGBに変換して <paramref name="f"/> を適用し、sRGB に戻す。</summary>
        private static double ScaleLinear(double srgb, Func<double, double> f)
        {
            var linear = srgb <= 0.04045 ? srgb / 12.92 : Math.Pow((srgb + 0.055) / 1.055, 2.4);
            var result = Clamp01(f(linear));
            return result <= 0.0031308 ? result * 12.92 : (1.055 * Math.Pow(result, 1.0 / 2.4)) - 0.055;
        }

        /// <summary>HSL の輝度に <paramref name="f"/> を適用する(<c>lumMod</c>/<c>lumOff</c>)。</summary>
        private static void ModifyLuminance(ref double r, ref double g, ref double b, Func<double, double> f)
        {
            var max = Math.Max(r, Math.Max(g, b));
            var min = Math.Min(r, Math.Min(g, b));
            var l = (max + min) / 2.0;
            double h = 0.0, s = 0.0;

            if (max > min)
            {
                var d = max - min;
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
                if (max == r)
                {
                    h = ((g - b) / d) + (g < b ? 6.0 : 0.0);
                }
                else if (max == g)
                {
                    h = ((b - r) / d) + 2.0;
                }
                else
                {
                    h = ((r - g) / d) + 4.0;
                }

                h /= 6.0;
            }

            l = Clamp01(f(l));

            if (s == 0.0)
            {
                r = g = b = l;
                return;
            }

            var q = l < 0.5 ? l * (1.0 + s) : l + s - (l * s);
            var p = (2.0 * l) - q;
            r = HueToRgb(p, q, h + (1.0 / 3.0));
            g = HueToRgb(p, q, h);
            b = HueToRgb(p, q, h - (1.0 / 3.0));
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0.0)
            {
                t += 1.0;
            }

            if (t > 1.0)
            {
                t -= 1.0;
            }

            if (t < 1.0 / 6.0)
            {
                return p + ((q - p) * 6.0 * t);
            }

            if (t < 1.0 / 2.0)
            {
                return q;
            }

            if (t < 2.0 / 3.0)
            {
                return p + ((q - p) * ((2.0 / 3.0) - t) * 6.0);
            }

            return p;
        }
    }
}
