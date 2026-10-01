using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Utsushi;
using Utsushi.Layout;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.OpenXml;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.Substitution;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// 帳票定義なしモードの列幅換算に使う最大数字幅(要件12.3、タスク27.2)の、ファサード(<see cref="ReportPdfConverter"/>)経由の検証。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 入力には invoice のサンプルを複製し、<c>xl/styles.xml</c> の先頭のフォント(ブックの標準フォント)の名前だけを差し替えたブックを使う
    /// (OpenXml SDK に依存してよいのは Parsing だけなので、ZIP内のXMLを文字列置換して作る。
    /// <see cref="WithoutDefinitionConversionTests"/> と同じ方針)。
    /// </para>
    /// <para>
    /// invoice の列幅は A〜F = 4, 28, 8, 6, 12, 14(文字数単位)。印刷範囲 A1:F34 は拡大縮小なし(100%)で1ページに収まるため、
    /// 罫線の左端から右端までの幅は列幅の換算結果の合計に一致する。
    /// </para>
    /// </remarks>
    public sealed class WithoutDefinitionMaxDigitWidthTests : IDisposable
    {
        private static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

        private static readonly double[] InvoiceColumnWidths = { 4.0, 28.0, 8.0, 6.0, 12.0, 14.0 };

        private readonly string _directory;

        public WithoutDefinitionMaxDigitWidthTests()
        {
            _directory = Path.Combine(Path.GetTempPath(), "utsushi-mdw-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        public void Dispose()
        {
            if (Directory.Exists(_directory))
            {
                Directory.Delete(_directory, recursive: true);
            }
        }

        private static ReportPdfConverter CreateConverter(string definitionRoot)
        {
            var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());

            return new ReportPdfConverter(
                new OpenXmlWorkbookReader(),
                new FileSystemReportDefinitionRepository(definitionRoot),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                new SkiaPdfRenderer(new SkiaFontMetricsProvider(fontResolver)),
                fontResolver);
        }

        private ReportPdfConverter CreateConverter() => CreateConverter(TestPaths.SampleReportsRoot);

        /// <summary>invoice のテンプレートを複製し、標準フォント(<c>fonts</c> の先頭)の名前を差し替えたブックを作る。</summary>
        private string CreateInvoiceWithDefaultFont(string fontName)
        {
            var path = Path.Combine(_directory, "invoice-" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(TestPaths.SampleTemplate("invoice"), path);

            using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
            var entry = archive.GetEntry("xl/styles.xml")
                ?? throw new InvalidOperationException("invoice のテンプレートに styles.xml がありません。");

            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
            {
                xml = reader.ReadToEnd();
            }

            // 先頭の <font> が標準フォント。サイズ11であることも確かめておく(見積もりの表は11ptだけを持つ)。
            const string firstFont = "<x:font><x:sz val=\"11\" /><x:color rgb=\"FF000000\" /><x:name val=\"MS PGothic\" />";
            var index = xml.IndexOf("<x:font>", StringComparison.Ordinal);
            if (index < 0 || string.CompareOrdinal(xml, index, firstFont, 0, firstFont.Length) != 0)
            {
                throw new InvalidOperationException("invoice のテンプレートの標準フォントが想定(MS PGothic 11pt)と異なります。");
            }

            var replacedFont = firstFont.Replace("MS PGothic", fontName, StringComparison.Ordinal);
            xml = xml.Substring(0, index) + replacedFont + xml.Substring(index + firstFont.Length);

            entry.Delete();
            var replaced = archive.CreateEntry("xl/styles.xml");
            using (var writer = new StreamWriter(replaced.Open(), new UTF8Encoding(false)))
            {
                writer.Write(xml);
            }

            return path;
        }

        /// <summary>invoice の帳票定義を複製し、<c>maxDigitWidthPx</c> だけを差し替えた定義ルートを作る。</summary>
        private string CreateDefinitionRootWithMaxDigitWidth(double maxDigitWidthPx)
        {
            var root = Path.Combine(_directory, "definitions");
            var reportDirectory = Path.Combine(root, "invoice");
            Directory.CreateDirectory(reportDirectory);

            var json = File.ReadAllText(Path.Combine(TestPaths.SampleReportsRoot, "invoice", "definition.json"));
            const string original = "\"maxDigitWidthPx\": 7,";
            if (json.IndexOf(original, StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException("invoice の帳票定義の maxDigitWidthPx が想定(7)と異なります。");
            }

            json = json.Replace(
                original,
                "\"maxDigitWidthPx\": " + maxDigitWidthPx.ToString(CultureInfo.InvariantCulture) + ",",
                StringComparison.Ordinal);
            File.WriteAllText(Path.Combine(reportDirectory, "definition.json"), json);
            return root;
        }

        private static PagedLayout ComputeWithoutDefinition(ReportPdfConverter converter, string path, double? maxDigitWidthPx = null)
        {
            using var input = File.OpenRead(path);
            return converter.ComputeLayoutWithoutDefinition(input, documentName: "doc", maxDigitWidthPx: maxDigitWidthPx);
        }

        private static PagedLayout ComputeWithDefinition(ReportPdfConverter converter, string path)
        {
            using var input = File.OpenRead(path);
            return converter.ComputeLayout(
                "invoice",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社サンプル商事 御中",
                    ["InvoiceNo"] = "INV-0000",
                    ["TotalAmount"] = "¥0",
                });
        }

        /// <summary>
        /// 置換値に左右されない、ページ構成と罫線・背景・描画オブジェクトの配置を1行ずつにしたもの。
        /// 列幅が変われば罫線・図形の座標が変わる。
        /// </summary>
        private static List<string> PageStructure(PagedLayout layout)
        {
            var lines = new List<string>();
            foreach (var page in layout.Pages)
            {
                lines.Add(FormattableString.Invariant(
                    $"page {page.PageNumber} {page.WidthPt:0.##}x{page.HeightPt:0.##} scale={page.ScaleFactor:0.####} rows={page.RowRange} cols={page.ColumnRange}"));
                foreach (var command in page.Commands)
                {
                    switch (command)
                    {
                        case LineCommand line:
                            lines.Add(FormattableString.Invariant(
                                $"line {line.From.X:0.##},{line.From.Y:0.##}-{line.To.X:0.##},{line.To.Y:0.##}"));
                            break;
                        case FillRectCommand fill:
                            lines.Add(FormattableString.Invariant($"fill {fill.Rect}"));
                            break;
                        case ShapeCommand shape:
                            lines.Add(FormattableString.Invariant($"shape {shape.Rect}"));
                            break;
                        case ImageCommand image:
                            lines.Add(FormattableString.Invariant($"image {image.Rect}"));
                            break;
                        case ConnectorCommand connector:
                            lines.Add(FormattableString.Invariant($"connector {connector.Rect}"));
                            break;
                        case GroupCommand group:
                            lines.Add(FormattableString.Invariant($"group {group.Center.X:0.##},{group.Center.Y:0.##}"));
                            break;
                    }
                }
            }

            return lines;
        }

        /// <summary>罫線の左端から右端までの幅(pt)。</summary>
        private static double TableWidth(PagedLayout layout)
        {
            var page = Assert.Single(layout.Pages);
            Assert.Equal(1.0, page.ScaleFactor, 6);
            // 垂直の罫線の位置で測る(水平の罫線は、角の継ぎ目で端が垂直の罫線の線幅の半分だけ延びる)。
            var xs = page.Commands.OfType<LineCommand>().Where(l => l.From.X == l.To.X).Select(l => l.From.X).ToList();
            return xs.Max() - xs.Min();
        }

        private static double ExpectedInvoiceWidth(double maxDigitWidthPx) =>
            InvoiceColumnWidths.Sum(width => ExcelUnitConverter.ColumnWidthToPoints(width, maxDigitWidthPx));

        // -- 呼び出し元の指定 ------------------------------------------------------------------

        [Theory]
        [InlineData(6.0)]
        [InlineData(7.0)]
        [InlineData(8.0)]
        [InlineData(9.0)]
        public void 最大数字幅を指定するとその値で列幅が換算される(double maxDigitWidthPx)
        {
            using var converter = CreateConverter();

            var layout = ComputeWithoutDefinition(converter, TestPaths.SampleTemplate("invoice"), maxDigitWidthPx);

            Assert.Equal(ExpectedInvoiceWidth(maxDigitWidthPx), TableWidth(layout), 3);
        }

        [Fact]
        public void 最大数字幅の指定が違えば列幅が変わる()
        {
            using var converter = CreateConverter();
            var path = TestPaths.SampleTemplate("invoice");

            var narrow = ComputeWithoutDefinition(converter, path, 7.0);
            var wide = ComputeWithoutDefinition(converter, path, 9.0);

            Assert.True(
                TableWidth(wide) > TableWidth(narrow),
                $"最大数字幅9の表の幅({TableWidth(wide)})が7の表の幅({TableWidth(narrow)})より広いこと");
            Assert.NotEqual(PageStructure(narrow), PageStructure(wide));
        }

        [Fact]
        public void 指定した最大数字幅は標準フォントからの見積もりより優先される()
        {
            // 標準フォントは ＭＳ Ｐゴシック 11(見積もりは8)。7を指定すると Calibri 11(見積もりも7)と同じになる。
            using var converter = CreateConverter();
            var gothic = CreateInvoiceWithDefaultFont("ＭＳ Ｐゴシック");
            var calibri = CreateInvoiceWithDefaultFont("Calibri");

            var specified = ComputeWithoutDefinition(converter, gothic, 7.0);
            var estimated = ComputeWithoutDefinition(converter, calibri);

            Assert.Equal(ExpectedInvoiceWidth(7.0), TableWidth(specified), 3);
            Assert.Equal(PageStructure(estimated), PageStructure(specified));
        }

        // -- 標準フォントからの見積もり -------------------------------------------------------------

        [Theory]
        [InlineData("ＭＳ Ｐゴシック")]
        [InlineData("MS PGothic")]
        [InlineData("游ゴシック")]
        public void 指定しなければ標準フォントから見積もった値になり定義ありで8を指定した場合と同じページ構成になる(string fontName)
        {
            var path = CreateInvoiceWithDefaultFont(fontName);
            using var converter = CreateConverter(CreateDefinitionRootWithMaxDigitWidth(8.0));

            var withoutDefinition = ComputeWithoutDefinition(converter, path);
            var withDefinition = ComputeWithDefinition(converter, path);

            Assert.Equal(ExpectedInvoiceWidth(8.0), TableWidth(withoutDefinition), 3);
            Assert.Equal(PageStructure(withDefinition), PageStructure(withoutDefinition));
        }

        [Fact]
        public void 標準フォントが見積もりの表と違えば定義ありで8を指定した場合と異なるページ構成になる()
        {
            // 上のテストが偶然一致していないこと(比較が列幅の違いを検出できること)の確認。
            var path = CreateInvoiceWithDefaultFont("Calibri");
            using var converter = CreateConverter(CreateDefinitionRootWithMaxDigitWidth(8.0));

            var withoutDefinition = ComputeWithoutDefinition(converter, path);
            var withDefinition = ComputeWithDefinition(converter, path);

            Assert.Equal(ExpectedInvoiceWidth(7.0), TableWidth(withoutDefinition), 3);
            Assert.NotEqual(PageStructure(withDefinition), PageStructure(withoutDefinition));
        }

        [Theory]
        [InlineData("Calibri")]
        [InlineData("Arial")]
        [InlineData("メイリオ")]
        public void 表に無い標準フォントやCalibriは7で換算され定義ありで7を指定した場合と同じページ構成になる(string fontName)
        {
            var path = CreateInvoiceWithDefaultFont(fontName);
            using var converter = CreateConverter(CreateDefinitionRootWithMaxDigitWidth(7.0));

            var withoutDefinition = ComputeWithoutDefinition(converter, path);
            var withDefinition = ComputeWithDefinition(converter, path);

            Assert.Equal(ExpectedInvoiceWidth(7.0), TableWidth(withoutDefinition), 3);
            Assert.Equal(PageStructure(withDefinition), PageStructure(withoutDefinition));
        }

        // -- 不正な値 ---------------------------------------------------------------------------

        public static IEnumerable<object[]> InvalidMaxDigitWidths() => new[]
        {
            new object[] { 0.0 },
            new object[] { -1.0 },
            new object[] { double.NaN },
            new object[] { double.PositiveInfinity },
            new object[] { double.NegativeInfinity },
        };

        [Theory]
        [MemberData(nameof(InvalidMaxDigitWidths))]
        public void ComputeLayoutWithoutDefinitionは不正な最大数字幅でArgumentOutOfRangeException(double maxDigitWidthPx)
        {
            using var converter = CreateConverter();
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<ArgumentOutOfRangeException>(
                () => converter.ComputeLayoutWithoutDefinition(input, maxDigitWidthPx: maxDigitWidthPx));

            Assert.Equal("maxDigitWidthPx", ex.ParamName);
        }

        [Theory]
        [MemberData(nameof(InvalidMaxDigitWidths))]
        public void ConvertWithoutDefinitionは不正な最大数字幅でArgumentOutOfRangeExceptionになり何も書き出さない(double maxDigitWidthPx)
        {
            using var converter = CreateConverter();
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var output = new MemoryStream();

            var ex = Assert.Throws<ArgumentOutOfRangeException>(
                () => converter.ConvertWithoutDefinition(input, output, maxDigitWidthPx: maxDigitWidthPx));

            Assert.Equal("maxDigitWidthPx", ex.ParamName);
            Assert.Equal(0, output.Length);
        }

        [Theory]
        [MemberData(nameof(InvalidMaxDigitWidths))]
        public void ConvertFileWithoutDefinitionは不正な最大数字幅でArgumentOutOfRangeExceptionになりファイルを出力しない(double maxDigitWidthPx)
        {
            using var converter = CreateConverter();
            var outputPath = Path.Combine(_directory, "out.pdf");

            var ex = Assert.Throws<ArgumentOutOfRangeException>(
                () => converter.ConvertFileWithoutDefinition(
                    TestPaths.SampleTemplate("invoice"), outputPath, maxDigitWidthPx: maxDigitWidthPx));

            Assert.Equal("maxDigitWidthPx", ex.ParamName);
            Assert.False(File.Exists(outputPath));
        }
    }
}
