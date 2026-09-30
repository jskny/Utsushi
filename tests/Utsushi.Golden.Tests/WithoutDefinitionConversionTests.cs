using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Utsushi;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
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
    /// 帳票定義なしモード(要件12、タスク24)のファサード(<see cref="ReportPdfConverter"/>)経由の検証。
    /// </summary>
    /// <remarks>
    /// <see cref="ReportConversionGoldenTests"/> と同じく、レイアウトの比較には決定的な
    /// <see cref="ApproximateFontMetricsProvider"/> を使い、描画には同梱フォントへのフォールバックを使う。
    /// </remarks>
    public sealed class WithoutDefinitionConversionTests
    {
        private static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

        /// <summary>
        /// 決定的なフォントメトリクスでレイアウトするコンバータ。<paramref name="renderer"/> が null なら
        /// <see cref="SkiaPdfRenderer"/>(同梱の日本語フォントへフォールバック)で描画する。
        /// 帳票定義リポジトリは、定義ありの変換と比較できるようサンプルのルートを指す。
        /// </summary>
        private static ReportPdfConverter CreateConverter(IPdfRenderer? renderer = null)
        {
            var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            var skiaMetrics = new SkiaFontMetricsProvider(fontResolver);

            return new ReportPdfConverter(
                new OpenXmlWorkbookReader(),
                new FileSystemReportDefinitionRepository(TestPaths.SampleReportsRoot),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                renderer ?? new SkiaPdfRenderer(skiaMetrics),
                fontResolver);
        }

        private static List<string> Texts(PagedLayout layout) =>
            layout.Pages.SelectMany(p => p.Commands.OfType<TextCommand>()).Select(t => t.Text).ToList();

        private static int CountShapes(IEnumerable<DrawCommand> commands) =>
            commands.Sum(c => c switch
            {
                ShapeCommand => 1,
                GroupCommand group => CountShapes(group.Children),
                _ => 0,
            });

        private static int CountShapes(PagedLayout layout) => layout.Pages.Sum(p => CountShapes(p.Commands));

        [Fact]
        public void 定義なしでアクティブシートを変換し文書名が帳票コードになる()
        {
            using var converter = CreateConverter();
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var layout = converter.ComputeLayoutWithoutDefinition(input, documentName: "請求書_2026年4月");

            Assert.Equal("請求書_2026年4月", layout.ReportCode);
            Assert.Equal("請求書", layout.SheetName);
            Assert.Equal(1, layout.PageCount);

            // テンプレートの値がそのまま描画される(置換キーは持たない)。
            var texts = Texts(layout);
            Assert.Contains("件名: サンプル案件", texts);
            Assert.Contains("請求番号", texts);
        }

        [Fact]
        public void 文書名を省略すると帳票コードは空文字列になる()
        {
            using var converter = CreateConverter();
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var layout = converter.ComputeLayoutWithoutDefinition(input);

            Assert.Equal(string.Empty, layout.ReportCode);
            Assert.Equal("請求書", layout.SheetName);
        }

        [Fact]
        public void 定義なしでも定義ありと同じページ構成と図形になる()
        {
            // 定義なしモードは同じ Substitution → Layout を通るため、置換値以外は定義ありと同じ結果になるべき。
            using var converter = CreateConverter();

            PagedLayout withDefinition;
            using (var input = File.OpenRead(TestPaths.SampleTemplate("invoice")))
            {
                withDefinition = converter.ComputeLayout(
                    "invoice",
                    input,
                    new Dictionary<string, string>
                    {
                        ["CustomerName"] = "株式会社テスト製作所 御中",
                        ["InvoiceNo"] = "INV-2026-0417",
                        ["TotalAmount"] = "¥2,153,800",
                    });
            }

            PagedLayout withoutDefinition;
            using (var input = File.OpenRead(TestPaths.SampleTemplate("invoice")))
            {
                withoutDefinition = converter.ComputeLayoutWithoutDefinition(input, documentName: "invoice");
            }

            Assert.Equal(withDefinition.PageCount, withoutDefinition.PageCount);
            Assert.Equal(withDefinition.Pages[0].WidthPt, withoutDefinition.Pages[0].WidthPt);
            Assert.Equal(withDefinition.Pages[0].HeightPt, withoutDefinition.Pages[0].HeightPt);
            Assert.Equal(withDefinition.Pages[0].RowRange, withoutDefinition.Pages[0].RowRange);
            Assert.Equal(withDefinition.Pages[0].ColumnRange, withoutDefinition.Pages[0].ColumnRange);
            Assert.Equal(CountShapes(withDefinition), CountShapes(withoutDefinition));
            Assert.True(CountShapes(withoutDefinition) > 0, "サンプルの図形が描画されていること");
        }

        [Fact]
        public void セル番地直接指定で結合セルのアンカーを上書きできる()
        {
            // A5:C5(件名)は結合セル。アンカーの A5 を指定する。
            using var converter = CreateConverter();
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var layout = converter.ComputeLayoutWithoutDefinition(
                input,
                new Dictionary<string, string> { ["A5"] = "件名: 定義なしで上書き", ["B12"] = "臨時オプション対応費" },
                "doc");

            var texts = Texts(layout);
            Assert.Contains("件名: 定義なしで上書き", texts);
            Assert.DoesNotContain("件名: サンプル案件", texts);
            Assert.Contains("臨時オプション対応費", texts);
            Assert.DoesNotContain("サンプル設計費", texts);
        }

        [Fact]
        public void セル番地直接指定で結合セルの非アンカーを指定するとエラーになる()
        {
            using var converter = CreateConverter();
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<NonAnchorMergedCellOverrideException>(() => converter.ComputeLayoutWithoutDefinition(
                input,
                new Dictionary<string, string> { ["B5"] = "見えなくなる値" },
                "doc"));

            Assert.Equal("doc", ex.ReportCode);
        }

        [Fact]
        public void 定義ありならunsupportedElementsがerrorで失敗する入力でも定義なしでは無視して変換できる()
        {
            // invoice の定義は unsupportedElements: "error"。トップレベルの星(star6)を非対応プリセットに差し替えたブックを作る。
            var path = CreateInvoiceWithUnsupportedShape();
            try
            {
                using var converter = CreateConverter();

                using (var input = File.OpenRead(path))
                {
                    var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => converter.ComputeLayout(
                        "invoice",
                        input,
                        new Dictionary<string, string>
                        {
                            ["CustomerName"] = "x",
                            ["InvoiceNo"] = "y",
                            ["TotalAmount"] = "z",
                        }));
                    Assert.Equal("UnsupportedShapePreset", ex.ElementKind);
                }

                int originalShapes;
                using (var input = File.OpenRead(TestPaths.SampleTemplate("invoice")))
                {
                    originalShapes = CountShapes(converter.ComputeLayoutWithoutDefinition(input));
                }

                using (var input = File.OpenRead(path))
                {
                    var layout = converter.ComputeLayoutWithoutDefinition(input, documentName: "doc");

                    Assert.Equal(1, layout.PageCount);
                    // 非対応の図形1つだけが描画されない。
                    Assert.Equal(originalShapes - 1, CountShapes(layout));
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ConvertWithoutDefinitionはPDFをストリームへ書き出す()
        {
            using var converter = CreateConverter();
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var output = new MemoryStream();

            converter.ConvertWithoutDefinition(input, output, documentName: "doc");

            var head = Encoding.ASCII.GetString(output.ToArray(), 0, 5);
            Assert.Equal("%PDF-", head);
        }

        [Fact]
        public void ConvertWithoutDefinitionは描画に渡すレイアウトの帳票コードが文書名になる()
        {
            var renderer = new CapturingPdfRenderer();
            using var converter = CreateConverter(renderer);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var output = new MemoryStream();

            converter.ConvertWithoutDefinition(input, output, documentName: "文書A");

            Assert.Equal("文書A", renderer.Captured!.ReportCode);
            Assert.Equal(CapturingPdfRenderer.Content, Encoding.UTF8.GetString(output.ToArray()));
        }

        [Fact]
        public void ConvertFileWithoutDefinitionは拡張子を除いたファイル名を文書名にしてファイルを出力する()
        {
            var directory = CreateTempDirectory();
            try
            {
                var inputPath = Path.Combine(directory, "請求書_2026年4月.xlsx");
                File.Copy(TestPaths.SampleTemplate("invoice"), inputPath);
                var outputPath = Path.Combine(directory, "out", "result.pdf");

                var renderer = new CapturingPdfRenderer();
                using var converter = CreateConverter(renderer);

                converter.ConvertFileWithoutDefinition(
                    inputPath, outputPath, new Dictionary<string, string> { ["A5"] = "件名: ファイル版" });

                Assert.Equal("請求書_2026年4月", renderer.Captured!.ReportCode);
                Assert.Equal("請求書", renderer.Captured.SheetName);
                Assert.Contains("件名: ファイル版", Texts(renderer.Captured));
                Assert.Equal(CapturingPdfRenderer.Content, File.ReadAllText(outputPath));
                Assert.False(File.Exists(outputPath + ".utsushi-tmp"), "一時ファイルが残ってはいけない");
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void ConvertFileWithoutDefinitionはSkiaPdfRendererでPDFファイルを出力する()
        {
            var directory = CreateTempDirectory();
            try
            {
                var outputPath = Path.Combine(directory, "invoice.pdf");
                using var converter = CreateConverter();

                converter.ConvertFileWithoutDefinition(TestPaths.SampleTemplate("invoice"), outputPath);

                var bytes = File.ReadAllBytes(outputPath);
                Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
                Assert.False(File.Exists(outputPath + ".utsushi-tmp"), "一時ファイルが残ってはいけない");
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void ConvertFileWithoutDefinitionは入力ファイルが存在しない場合UtsushiExceptionに変換し出力しない()
        {
            var directory = CreateTempDirectory();
            try
            {
                var missingPath = Path.Combine(directory, "存在しない帳票.xlsx");
                var outputPath = Path.Combine(directory, "out.pdf");
                using var converter = CreateConverter(new CapturingPdfRenderer());

                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => converter.ConvertFileWithoutDefinition(missingPath, outputPath));

                Assert.Equal(ProcessingStage.Parsing, ex.Stage);
                Assert.Equal("存在しない帳票", ex.ReportCode);
                Assert.False(File.Exists(outputPath));
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void 定義ルートなしのCreateDefaultで定義なし変換ができる()
        {
            using var converter = ReportPdfConverter.CreateDefault(FontResolverOptions.AllowFallback());
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var output = new MemoryStream();

            converter.ConvertWithoutDefinition(input, output, documentName: "doc");

            Assert.Equal("%PDF-", Encoding.ASCII.GetString(output.ToArray(), 0, 5));
        }

        [Fact]
        public void 定義ルートなしのCreateDefaultでComputeLayoutに帳票コードを渡すと帳票定義が見つからないエラーになる()
        {
            using var converter = ReportPdfConverter.CreateDefault(FontResolverOptions.AllowFallback());
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<ReportDefinitionNotFoundException>(
                () => converter.ComputeLayout("invoice", input, new Dictionary<string, string>()));

            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal(ProcessingStage.ReportDefinition, ex.Stage);
        }

        [Fact]
        public void 定義ルートなしのCreateDefaultでConvertToFileに帳票コードを渡すと帳票定義が見つからないエラーになり出力しない()
        {
            var directory = CreateTempDirectory();
            try
            {
                var outputPath = Path.Combine(directory, "invoice.pdf");
                using var converter = ReportPdfConverter.CreateDefault(FontResolverOptions.AllowFallback());

                var ex = Assert.Throws<ReportDefinitionNotFoundException>(() => converter.ConvertToFile(
                    "invoice",
                    TestPaths.SampleTemplate("invoice"),
                    new Dictionary<string, string>
                    {
                        ["CustomerName"] = "x",
                        ["InvoiceNo"] = "y",
                        ["TotalAmount"] = "z",
                    },
                    outputPath));

                Assert.Equal("invoice", ex.ReportCode);
                Assert.False(File.Exists(outputPath));
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void ComputeLayoutWithoutDefinitionはnullのストリームでArgumentNullException()
        {
            using var converter = CreateConverter(new CapturingPdfRenderer());

            Assert.Throws<ArgumentNullException>(() => converter.ComputeLayoutWithoutDefinition(null!));
        }

        [Fact]
        public void ConvertWithoutDefinitionはnullの出力先でArgumentNullException()
        {
            using var converter = CreateConverter(new CapturingPdfRenderer());
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            Assert.Throws<ArgumentNullException>(() => converter.ConvertWithoutDefinition(input, null!));
        }

        private static string CreateTempDirectory()
        {
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            return directory;
        }

        /// <summary>
        /// invoice のテンプレートを複製し、トップレベルの星(<c>prst="star6"</c>)を非対応プリセットに差し替えたブックを作る。
        /// </summary>
        /// <remarks>
        /// OpenXml SDK に依存してよいのは Parsing だけ(<c>.kiro/steering/structure.md</c>)なので、
        /// ZIP内のXMLを文字列置換して作る。
        /// </remarks>
        private static string CreateInvoiceWithUnsupportedShape()
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-unsupported-" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(TestPaths.SampleTemplate("invoice"), path);

            using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
            var entry = archive.GetEntry("xl/drawings/drawing1.xml")
                ?? throw new InvalidOperationException("invoice のテンプレートに描画パートがありません。");

            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
            {
                xml = reader.ReadToEnd();
            }

            const string original = "prst=\"star6\"";
            if (xml.IndexOf(original, StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException("invoice のテンプレートに星(star6)の図形がありません。");
            }

            xml = xml.Replace(original, "prst=\"flowChartPreparation\"", StringComparison.Ordinal);

            entry.Delete();
            var replaced = archive.CreateEntry("xl/drawings/drawing1.xml");
            using (var writer = new StreamWriter(replaced.Open(), new UTF8Encoding(false)))
            {
                writer.Write(xml);
            }

            return path;
        }

        /// <summary>描画に渡されたレイアウトを捕まえる、テスト用の最小の <see cref="IPdfRenderer"/>。</summary>
        private sealed class CapturingPdfRenderer : IPdfRenderer
        {
            public const string Content = "fake-pdf-content";

            public PagedLayout? Captured { get; private set; }

            public void Render(PagedLayout layout, Stream output)
            {
                Captured = layout;
                var bytes = Encoding.UTF8.GetBytes(Content);
                output.Write(bytes, 0, bytes.Length);
            }
        }
    }
}
