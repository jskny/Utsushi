using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Utsushi.Core;
using Utsushi.Parsing.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 図形・接続線の色(DrawingML の <c>a:srgbClr</c>/<c>a:schemeClr</c>/<c>a:sysClr</c> と色の修飾)の解決と、
    /// テーマの線の書式の太さの検証(要件10.15, 10.16。design.md「テーマの色・スタイル参照」)。
    /// </summary>
    /// <remarks>
    /// 期待値のうち Excel の UI に現れる色(「アクセント1、黒+基本色25%」= 2F5597、「白+基本色40%」= 8FAADC、
    /// 既定の図形の枠線 = 2F528F)は、Excel が保存した値と一致することを確認済みの値を使う。
    /// </remarks>
    public sealed class DrawingColorResolverTests
    {
        private const string Ns = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";

        private static A.SolidFill Fill(string colorXml) => new("<a:solidFill " + Ns + ">" + colorXml + "</a:solidFill>");

        private static ArgbColor Rgb(string hex)
        {
            Assert.True(ArgbColor.TryParseHex(hex, out var color));
            return color;
        }

        private static ArgbColor Resolve(string colorXml, ArgbColor? placeholder = null, DrawingColorResolver? resolver = null)
        {
            Assert.True((resolver ?? DrawingColorResolver.Default).TryResolve(Fill(colorXml), placeholder, out var color));
            return color;
        }

        // -- 色の種類 -------------------------------------------------

        [Theory]
        [InlineData("<a:srgbClr val=\"1F4E8C\"/>", "1F4E8C")]
        [InlineData("<a:schemeClr val=\"accent1\"/>", "4472C4")]
        [InlineData("<a:schemeClr val=\"accent2\"/>", "ED7D31")]
        [InlineData("<a:schemeClr val=\"accent6\"/>", "70AD47")]
        [InlineData("<a:schemeClr val=\"dk1\"/>", "000000")]
        [InlineData("<a:schemeClr val=\"lt1\"/>", "FFFFFF")]
        [InlineData("<a:schemeClr val=\"dk2\"/>", "44546A")]
        [InlineData("<a:schemeClr val=\"lt2\"/>", "E7E6E6")]
        [InlineData("<a:schemeClr val=\"hlink\"/>", "0563C1")]
        [InlineData("<a:schemeClr val=\"folHlink\"/>", "954F72")]
        [InlineData("<a:sysClr val=\"windowText\" lastClr=\"123456\"/>", "123456")]
        [InlineData("<a:sysClr val=\"window\"/>", "FFFFFF")]
        [InlineData("<a:sysClr val=\"windowText\"/>", "000000")]
        public void 色の種類ごとに色を解決する(string colorXml, string expectedHex)
        {
            // 要件10.15。sysClr は lastClr を優先し、無ければ window=白・それ以外=黒。
            Assert.Equal(Rgb(expectedHex), Resolve(colorXml));
        }

        [Theory]
        [InlineData("tx1", "dk1")]
        [InlineData("bg1", "lt1")]
        [InlineData("tx2", "dk2")]
        [InlineData("bg2", "lt2")]
        public void tx1とbg1とtx2とbg2はdk1とlt1とdk2とlt2の別名として解決する(string alias, string target)
        {
            Assert.Equal(
                Resolve("<a:schemeClr val=\"" + target + "\"/>"),
                Resolve("<a:schemeClr val=\"" + alias + "\"/>"));
        }

        [Fact]
        public void phClrは呼び出し側が渡した色に修飾をかけて解決する()
        {
            var placeholder = Rgb("4472C4");

            Assert.Equal(placeholder, Resolve("<a:schemeClr val=\"phClr\"/>", placeholder));
            Assert.Equal(Rgb("2F528F"), Resolve("<a:schemeClr val=\"phClr\"><a:shade val=\"50000\"/></a:schemeClr>", placeholder));
        }

        [Fact]
        public void phClrに差し込む色が無ければ解決できない()
        {
            Assert.False(DrawingColorResolver.Default.TryResolve(Fill("<a:schemeClr val=\"phClr\"/>"), null, out _));
        }

        [Theory]
        [InlineData("<a:prstClr val=\"red\"/>")]
        [InlineData("<a:scrgbClr r=\"100000\" g=\"0\" b=\"0\"/>")]
        [InlineData("<a:hslClr hue=\"0\" sat=\"100000\" lum=\"50000\"/>")]
        [InlineData("<a:srgbClr val=\"XYZ\"/>")]
        [InlineData("")]
        public void 未対応の色の指定や色の要素が無い場合は解決できない(string colorXml)
        {
            Assert.False(DrawingColorResolver.Default.TryResolve(Fill(colorXml), null, out _));
        }

        [Fact]
        public void 色を持つ要素がnullなら解決できない()
        {
            Assert.False(DrawingColorResolver.Default.TryResolve(null, null, out _));
        }

        [Fact]
        public void 塗りつぶし以外の要素でも子の色の要素から解決する()
        {
            // スタイル参照(lnRef/fillRef/fontRef)やグラデーションの分岐点(gs)も、色の要素を子に持つ。
            var lineRef = new A.LineReference("<a:lnRef " + Ns + " idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef>");
            var stop = new A.GradientStop("<a:gs " + Ns + " pos=\"0\"><a:schemeClr val=\"accent2\"/></a:gs>");

            Assert.True(DrawingColorResolver.Default.TryResolve(lineRef, null, out var lineColor));
            Assert.Equal(Rgb("2F528F"), lineColor);
            Assert.True(DrawingColorResolver.Default.TryResolve(stop, null, out var stopColor));
            Assert.Equal(Rgb("ED7D31"), stopColor);
        }

        // -- 色の修飾 -------------------------------------------------

        [Theory]
        [InlineData("<a:lumMod val=\"75000\"/>", "2F5597")] // 黒+基本色25%
        [InlineData("<a:lumMod val=\"50000\"/>", "203864")] // 黒+基本色50%
        [InlineData("<a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/>", "8FAADC")] // 白+基本色40%
        [InlineData("<a:lumMod val=\"20000\"/><a:lumOff val=\"80000\"/>", "DAE3F3")] // 白+基本色80%
        public void lumModとlumOffはHSLの輝度を変える(string modifiers, string expectedHex)
        {
            Assert.Equal(Rgb(expectedHex), Resolve("<a:schemeClr val=\"accent1\">" + modifiers + "</a:schemeClr>"));
        }

        [Fact]
        public void 無彩色のlumModは各成分をそのまま縮める()
        {
            Assert.Equal(Rgb("404040"), Resolve("<a:srgbClr val=\"808080\"><a:lumMod val=\"50000\"/></a:srgbClr>"));
        }

        [Theory]
        [InlineData("4472C4", "<a:shade val=\"50000\"/>", "2F528F")] // sRGBのまま半分にすると 223962 になる
        [InlineData("FF0000", "<a:shade val=\"50000\"/>", "BC0000")] // 同 800000
        [InlineData("4472C4", "<a:tint val=\"50000\"/>", "C0C9E4")] // 同 A2B8E2
        [InlineData("000000", "<a:tint val=\"50000\"/>", "BCBCBC")] // 同 808080
        [InlineData("4472C4", "<a:shade val=\"100000\"/>", "4472C4")]
        [InlineData("4472C4", "<a:tint val=\"100000\"/>", "4472C4")]
        [InlineData("4472C4", "<a:shade val=\"0\"/>", "000000")]
        [InlineData("4472C4", "<a:tint val=\"0\"/>", "FFFFFF")]
        public void shadeとtintは線形RGBで黒と白へ補間する(string baseHex, string modifier, string expectedHex)
        {
            Assert.Equal(Rgb(expectedHex), Resolve("<a:srgbClr val=\"" + baseHex + "\">" + modifier + "</a:srgbClr>"));
        }

        [Theory]
        [InlineData("50000", 0x80)]
        [InlineData("0", 0x00)]
        [InlineData("100000", 0xFF)]
        [InlineData("250000", 0xFF)] // 範囲外は丸める
        public void alphaは不透明度になり色の成分は変えない(string alpha, byte expectedAlpha)
        {
            var color = Resolve("<a:srgbClr val=\"1F4E8C\"><a:alpha val=\"" + alpha + "\"/></a:srgbClr>");

            Assert.Equal(new ArgbColor(expectedAlpha, 0x1F, 0x4E, 0x8C), color);
        }

        [Fact]
        public void 修飾は文書の順に適用する()
        {
            // lumOff を先にかけると白に張り付き、その後の lumMod で灰色になる(順序を入れ替えると結果が変わる)。
            var offThenMod = Resolve("<a:srgbClr val=\"4472C4\"><a:lumOff val=\"100000\"/><a:lumMod val=\"50000\"/></a:srgbClr>");
            var modThenOff = Resolve("<a:srgbClr val=\"4472C4\"><a:lumMod val=\"50000\"/><a:lumOff val=\"100000\"/></a:srgbClr>");

            Assert.Equal(Rgb("808080"), offThenMod);
            Assert.Equal(Rgb("FFFFFF"), modThenOff);
        }

        [Fact]
        public void 修飾は16個までを適用し17個目以降は無視する()
        {
            // 信頼できない入力への安全弁(MaxColorModifiers)。無害な lumMod 100% を並べて境界を確かめる。
            var noOps = new StringBuilder();
            for (var i = 0; i < 15; i++)
            {
                noOps.Append("<a:lumMod val=\"100000\"/>");
            }

            var sixteenth = Resolve("<a:srgbClr val=\"1F4E8C\">" + noOps + "<a:alpha val=\"0\"/></a:srgbClr>");
            var seventeenth = Resolve("<a:srgbClr val=\"1F4E8C\">" + noOps + "<a:lumMod val=\"100000\"/><a:alpha val=\"0\"/></a:srgbClr>");

            Assert.Equal(0x00, sixteenth.A);
            Assert.Equal(0xFF, seventeenth.A);
            Assert.Equal(Rgb("1F4E8C"), seventeenth);
        }

        [Fact]
        public void 未対応の修飾は無視する()
        {
            Assert.Equal(Rgb("1F4E8C"), Resolve("<a:srgbClr val=\"1F4E8C\"><a:satMod val=\"200000\"/><a:hueOff val=\"100000\"/></a:srgbClr>"));
        }

        // -- テーマ -------------------------------------------------

        [Fact]
        public void テーマの無いブックではOffice既定の配色と既定の線の太さになる()
        {
            WithWorkbookPart(themeXml: null, workbookPart =>
            {
                var resolver = DrawingColorResolver.Create(workbookPart);

                Assert.Same(DrawingColorResolver.Default, resolver);
                Assert.Equal(Rgb("4472C4"), Resolve("<a:schemeClr val=\"accent1\"/>", resolver: resolver));
                // Office 既定テーマの lnStyleLst(6350/12700/19050 EMU)と同じ太さ。
                Assert.Equal(0.5, resolver.GetLineStyleWidthPt(1), 6);
                Assert.Equal(1.0, resolver.GetLineStyleWidthPt(2), 6);
                Assert.Equal(1.5, resolver.GetLineStyleWidthPt(3), 6);
                Assert.Equal(DrawingColorResolver.DefaultLineStyleWidthPt, resolver.GetLineStyleWidthPt(4), 6);
            });
        }

        [Fact]
        public void テーマの配色を使う()
        {
            var theme = DrawingStyleWorkbookFixtures.ThemeXml(accent1: "112233", dark1: "101010", 6350);
            WithWorkbookPart(theme, workbookPart =>
            {
                var resolver = DrawingColorResolver.Create(workbookPart);

                Assert.Equal(Rgb("112233"), Resolve("<a:schemeClr val=\"accent1\"/>", resolver: resolver));
                Assert.Equal(Rgb("101010"), Resolve("<a:schemeClr val=\"tx1\"/>", resolver: resolver));
                Assert.Equal(Rgb("FFFFFF"), Resolve("<a:schemeClr val=\"bg1\"/>", resolver: resolver)); // sysClr の lastClr
                Assert.Equal(Rgb("ED7D31"), Resolve("<a:schemeClr val=\"accent2\"/>", resolver: resolver));
            });
        }

        [Theory]
        [InlineData(1u, 0.5)]
        [InlineData(2u, 1.0)]
        [InlineData(3u, 1.5)]
        [InlineData(0u, 0.75)] // 範囲外は既定値
        [InlineData(4u, 0.75)]
        [InlineData(uint.MaxValue, 0.75)]
        public void lnRefの番号はテーマのlnStyleLstの線の太さを指す(uint index, double expectedPt)
        {
            WithWorkbookPart(DrawingStyleWorkbookFixtures.OfficeThemeXml(), workbookPart =>
            {
                Assert.Equal(expectedPt, DrawingColorResolver.Create(workbookPart).GetLineStyleWidthPt(index), 6);
            });
        }

        [Fact]
        public void lnStyleLstの線に太さが無ければ既定値になる()
        {
            var theme = DrawingStyleWorkbookFixtures.ThemeXml(lineStyleWidthsEmu: new long?[] { null, 25400 });
            WithWorkbookPart(theme, workbookPart =>
            {
                var resolver = DrawingColorResolver.Create(workbookPart);

                Assert.Equal(0.75, resolver.GetLineStyleWidthPt(1), 6);
                Assert.Equal(2.0, resolver.GetLineStyleWidthPt(2), 6);
            });
        }

        [Fact]
        public void テーマが無い場合のlnRefの太さはOffice既定テーマの線の太さになる()
        {
            Assert.Equal(1.0, DrawingColorResolver.Default.GetLineStyleWidthPt(2), 6);
            Assert.Equal(0.75, DrawingColorResolver.DefaultLineStyleWidthPt);
        }

        private static void WithWorkbookPart(string? themeXml, System.Action<WorkbookPart> action)
        {
            using var stream = new MemoryStream();
            using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            if (themeXml is not null)
            {
                var themePart = workbookPart.AddNewPart<ThemePart>();
                using var themeStream = new MemoryStream(new UTF8Encoding(false).GetBytes(themeXml));
                themePart.FeedData(themeStream);
            }

            action(workbookPart);
        }
    }
}
