using System;
using System.Collections.Generic;
using System.IO;
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
    /// 登録済み帳票サンプルのエンドツーエンド(Parsing → Rendering)ゴールデンテスト(タスク8.2、要件8.3)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// レイアウトの比較には <see cref="ApproximateFontMetricsProvider"/> を使う。
    /// 実フォントのメトリクスは実行環境のフォント構成に依存し、CIとローカルでゴールデンが
    /// 一致しなくなるため、回帰検出の基準としては決定的な近似メトリクスを用いる。
    /// </para>
    /// <para>
    /// 実フォントを使った見た目の確認は「PDFが生成できること」までを自動テストの範囲とし、
    /// 最終的な見た目のレビューは人が行う前提とする。
    /// </para>
    /// </remarks>
    public sealed class ReportConversionGoldenTests
    {
        /// <summary>
        /// 決定的なフォントメトリクスを使い、PDF描画の手前までを実行するコンバータ。
        /// </summary>
        /// <summary>
        /// ヘッダー/フッターの &amp;D(日付)・&amp;T(時刻)を固定するための時刻。
        /// 実時刻のままではゴールデンファイルが毎日変わってしまう。
        /// </summary>
        private static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

        private static ReportPdfConverter CreateDeterministicConverter(out FontResolver fontResolver)
        {
            fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            var skiaMetrics = new SkiaFontMetricsProvider(fontResolver);

            return new ReportPdfConverter(
                new OpenXmlWorkbookReader(),
                new FileSystemReportDefinitionRepository(TestPaths.SampleReportsRoot),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                // CIには日本語フォントが無いため、字形欠落の検出(要件5.5)は無効にする。
                // ゴールデンテストの比較対象はレイアウト結果であり、PDFの字形ではない。
                new SkiaPdfRenderer(skiaMetrics, PdfRenderOptions.Default with { MissingGlyphs = MissingGlyphPolicy.Render }),
                fontResolver);
        }

        public static IEnumerable<object[]> RegisteredReports()
        {
            yield return new object[]
            {
            "invoice",
            new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 御中",
                ["InvoiceNo"] = "INV-2026-0417",
                ["Subject"] = "件名: ゴールデンテスト用案件",
                ["TotalAmount"] = "¥2,153,800",
                ["Remarks"] = "備考:\nゴールデンテストで固定した値です。",
            },
            };

            yield return new object[]
            {
            "receipt",
            new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 様",
                ["Amount"] = "¥2,153,800",
                ["Description"] = "サンプル案件の代金として",
                ["CopyCustomerName"] = "株式会社テスト製作所 様",
                ["CopyAmount"] = "¥2,153,800",
                ["CopyDescription"] = "サンプル案件の代金として",
            },
            };

            yield return new object[]
            {
            "delivery-note",
            new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 御中",
                ["DeliveryNo"] = "DN-2026-0088",
                ["DeliveryDate"] = "2026年4月15日",
            },
            };
        }

        [Theory]
        [MemberData(nameof(RegisteredReports))]
        public void 登録済み帳票のレイアウトが期待結果と一致する(string reportCode, Dictionary<string, string> values)
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate(reportCode));

            var layout = converter.ComputeLayout(reportCode, input, values);

            GoldenFile.Verify(reportCode, "layout.snapshot.txt", GoldenSnapshot.Create(layout));
        }

        [Theory]
        [MemberData(nameof(RegisteredReports))]
        public void 登録済み帳票からPDFを生成できる(string reportCode, Dictionary<string, string> values)
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate(reportCode));
            using var output = new MemoryStream();

            converter.Convert(reportCode, input, values, output);

            var bytes = output.ToArray();
            Assert.True(bytes.Length > 0, "PDFが出力されるはず");
            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        }

        [Fact]
        public void 請求書は1ページに収まる()
        {
            var layout = ComputeInvoiceLayout();

            Assert.Equal(1, layout.PageCount);
            Assert.Equal("A4", layout.Pages[0].Paper.Name);
        }

        [Fact]
        public void 納品書は複数ページになり印刷タイトルが各ページに繰り返される()
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("delivery-note"));

            var layout = converter.ComputeLayout(
                "delivery-note",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社テスト製作所 御中",
                    ["DeliveryNo"] = "DN-2026-0088",
                });

            Assert.True(layout.PageCount > 1, $"複数ページになるはず (実際: {layout.PageCount})");

            // 印刷タイトル行(1〜7行目)の内容が全ページに現れる(要件3.4)
            foreach (var page in layout.Pages)
            {
                var texts = page.Commands.OfType<TextCommand>().Select(t => t.Text).ToList();
                Assert.Contains("納品書", texts);
                Assert.Contains("品名", texts);
                Assert.Contains("株式会社テスト製作所 御中", texts);
            }

            // 本文はページをまたいで連続する
            var firstPageRows = layout.Pages[0].RowRange;
            var secondPageRows = layout.Pages[1].RowRange;
            Assert.Equal(firstPageRows.Last + 1, secondPageRows.First);
        }

        [Fact]
        public void 置換値が帳票に反映される()
        {
            var layout = ComputeInvoiceLayout();
            var texts = layout.Pages
                .SelectMany(p => p.Commands.OfType<TextCommand>())
                .Select(t => t.Text)
                .ToList();

            Assert.Contains("株式会社テスト製作所 御中", texts);
            Assert.Contains("INV-2026-0417", texts);
            Assert.Contains("¥2,153,800", texts);

            // テンプレートの既定値は残らない
            Assert.DoesNotContain("INV-0000", texts);
        }

        [Fact]
        public void セル番地直接指定で帳票定義に未登録のセルを上書きできる()
        {
            // B12(明細1行目の品名セル)は invoice の definition.json に置換キーとして
            // 登録されていない。置換キー方式(values)と併用できることも合わせて検証する。
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var layout = converter.ComputeLayout(
                "invoice",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社テスト製作所 御中",
                    ["InvoiceNo"] = "INV-2026-0417",
                    ["TotalAmount"] = "¥2,153,800",
                },
                new Dictionary<string, string> { ["B12"] = "臨時オプション対応費" });

            var texts = layout.Pages
                .SelectMany(p => p.Commands.OfType<TextCommand>())
                .Select(t => t.Text)
                .ToList();

            Assert.Contains("臨時オプション対応費", texts);
            Assert.DoesNotContain("サンプル設計費", texts);

            // 置換キー方式の値も引き続き反映される。
            Assert.Contains("株式会社テスト製作所 御中", texts);
            Assert.Contains("INV-2026-0417", texts);
        }

        [Fact]
        public void 結合セルの非アンカーを直接指定するとエラーになる()
        {
            // A3:C3(CustomerNameのセル)は結合セルであり、B3・C3はアンカーではない。
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<NonAnchorMergedCellOverrideException>(() => converter.ComputeLayout(
                "invoice",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社テスト製作所 御中",
                    ["InvoiceNo"] = "INV-2026-0417",
                    ["TotalAmount"] = "¥2,153,800",
                },
                new Dictionary<string, string> { ["B3"] = "見えなくなる値" }));

            Assert.Equal(CellAddress.Parse("B3"), ex.CellAddress);
            Assert.Equal(CellAddress.Parse("A3"), ex.AnchorAddress);
            Assert.Equal("invoice", ex.ReportCode);
        }

        [Fact]
        public void 不正なセル番地を直接指定するとエラーになる()
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<InvalidCellOverrideAddressException>(() => converter.ComputeLayout(
                "invoice",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社テスト製作所 御中",
                    ["InvoiceNo"] = "INV-2026-0417",
                    ["TotalAmount"] = "¥2,153,800",
                },
                new Dictionary<string, string> { ["not-a-cell"] = "x" }));

            Assert.Equal("not-a-cell", ex.Address);
            Assert.Equal("invoice", ex.ReportCode);
        }

        [Fact]
        public void 必須の置換値が無ければ変換前に失敗する()
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<RequiredSubstitutionValueMissingException>(
                () => converter.ComputeLayout("invoice", input, new Dictionary<string, string>()));

            Assert.Equal("invoice", ex.ReportCode);
        }

        [Fact]
        public void 未登録の帳票コードは明確なエラーになる()
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            var ex = Assert.Throws<ReportDefinitionNotFoundException>(
                () => converter.ComputeLayout("未登録の帳票", input, new Dictionary<string, string>()));

            Assert.Equal("未登録の帳票", ex.ReportCode);
        }

        [Fact]
        public void 帳票定義と一致しないExcelを渡すと失敗する()
        {
            // 納品書のテンプレートを請求書の定義で読もうとするとシート名が一致しない(要件1.4)。
            // ファイル自体は正常な.xlsxのため、Parsing段階の障害ではなく帳票定義との構造不一致
            // (Stage=ReportDefinition)として報告されるべき(要件6.4のStageベースの分岐に対応)。
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("delivery-note"));

            var ex = Assert.Throws<ReportStructureMismatchException>(
                () => converter.ComputeLayout(
                    "invoice",
                    input,
                    new Dictionary<string, string>
                    {
                        ["CustomerName"] = "x",
                        ["InvoiceNo"] = "y",
                        ["TotalAmount"] = "z",
                    }));

            Assert.Equal(ProcessingStage.ReportDefinition, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.IsType<InvalidExcelFileException>(ex.InnerException);
        }

        [Fact]
        public void 複数の印刷範囲はそれぞれ別ページになり範囲外は出力されない()
        {
            // 領収書サンプルは「本紙」と「控え」を別々の印刷範囲として持つ(要件3.6)。
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("receipt"));

            var layout = converter.ComputeLayout(
                "receipt",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社テスト製作所 様",
                    ["Amount"] = "¥1,000",
                    ["CopyCustomerName"] = "株式会社テスト製作所 様",
                    ["CopyAmount"] = "¥1,000",
                });

            Assert.Equal(2, layout.PageCount);

            var firstPageTexts = layout.Pages[0].Commands.OfType<TextCommand>().Select(t => t.Text).ToList();
            var secondPageTexts = layout.Pages[1].Commands.OfType<TextCommand>().Select(t => t.Text).ToList();

            Assert.Contains("領 収 書", firstPageTexts);
            Assert.Contains("領 収 書(控)", secondPageTexts);
            Assert.DoesNotContain("領 収 書(控)", firstPageTexts);

            // 2つの印刷範囲の間にある行は、どちらの範囲にも含まれないため出力されない(要件3.1)。
            var allTexts = firstPageTexts.Concat(secondPageTexts).ToList();
            Assert.DoesNotContain(allTexts, t => t.Contains("印刷範囲外"));
        }

        [Fact]
        public void ヘッダーとフッターの書式コードが展開される()
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("delivery-note"));

            var layout = converter.ComputeLayout(
                "delivery-note",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社テスト製作所 御中",
                    ["DeliveryNo"] = "DN-2026-0088",
                });

            Assert.True(layout.PageCount >= 2);

            foreach (var page in layout.Pages)
            {
                var texts = page.Commands.OfType<TextCommand>().Select(t => t.Text).ToList();

                // ヘッダー &R&A → シート名
                Assert.Contains("納品書", texts);

                // フッター &L&D → 日付(固定した時刻)
                Assert.Contains("2026/04/20", texts);

                // フッター &C&P / &N ページ → ページ番号と総ページ数
                Assert.Contains($"{page.PageNumber} / {layout.PageCount} ページ", texts);

                // フッター &R&"MS PGothic,Bold"サンプル → 太字の指定が効いている
                var sample = texts.Contains("サンプル");
                Assert.True(sample, "右セクションの文字列が出力されるはず");
                Assert.Contains(
                    page.Commands.OfType<TextCommand>(),
                    t => t.Text == "サンプル" && t.Font.Bold);
            }
        }

        [Fact]
        public void 既定以外のIPdfRendererでもファイル出力は一時ファイル経由で確定する()
        {
            // ConvertToFileはSkiaPdfRenderer以外のIPdfRendererにも対応する(コンストラクタで差し替え可能)。
            // 要件5.4(失敗時に不完全なファイルを残さない)はレンダラーの実装によらず保証されるべき。
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = new ReportPdfConverter(
                new OpenXmlWorkbookReader(),
                new FileSystemReportDefinitionRepository(TestPaths.SampleReportsRoot),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                new FakePdfRenderer(),
                fontResolver);

            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            var outputPath = Path.Combine(directory, "invoice.pdf");

            try
            {
                converter.ConvertToFile(
                    "invoice",
                    TestPaths.SampleTemplate("invoice"),
                    new Dictionary<string, string>
                    {
                        ["CustomerName"] = "株式会社テスト製作所 御中",
                        ["InvoiceNo"] = "INV-2026-0417",
                        ["TotalAmount"] = "¥2,153,800",
                    },
                    outputPath);

                Assert.True(File.Exists(outputPath));
                Assert.Equal(FakePdfRenderer.Content, File.ReadAllText(outputPath));
                Assert.False(File.Exists(outputPath + ".utsushi-tmp"), "一時ファイルが残ってはいけない");
            }
            finally
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
        }

        private static Dictionary<string, string> InvoiceValues(string? remarks = null)
        {
            var values = new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 御中",
                ["InvoiceNo"] = "INV-2026-0417",
                ["TotalAmount"] = "¥2,153,800",
            };

            if (remarks is not null)
            {
                values["Remarks"] = remarks;
            }

            return values;
        }

        private static PagedLayout ComputeInvoice(IReadOnlyDictionary<string, string> values)
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            return converter.ComputeLayout("invoice", input, values);
        }

        [Fact]
        public void 差し込み値の制御文字はファサード経由でもSubstitution段階のエラーになる()
        {
            var values = InvoiceValues();
            values["InvoiceNo"] = "INV\t0001";

            var ex = Assert.Throws<InvalidSubstitutionValueException>(() => ComputeInvoice(values));

            Assert.Equal(ProcessingStage.Substitution, ex.Stage);
            Assert.Equal("invoice", ex.ReportCode);
            Assert.Equal(CellAddress.Parse("F3"), ex.CellAddress);
        }

        [Fact]
        public void 備考欄のCRLF区切りの住所はLF区切りと同じレイアウトになる()
        {
            static IEnumerable<string> RemarkLines(PagedLayout layout) =>
                layout.Pages.SelectMany(p => p.Commands.OfType<TextCommand>())
                    .Where(t => t.Text.Contains("丸の内") || t.Text.Contains("東京都") || t.Text.Contains("担当"))
                    .Select(t => t.Text);

            var lf = ComputeInvoice(InvoiceValues("東京都千代田区丸の内1-1-1\n丸の内ビル10F\n担当: 山田"));
            var crlf = ComputeInvoice(InvoiceValues("東京都千代田区丸の内1-1-1\r\n丸の内ビル10F\r\n担当: 山田"));

            Assert.Equal(new[] { "東京都千代田区丸の内1-1-1", "丸の内ビル10F", "担当: 山田" }, RemarkLines(lf));
            Assert.Equal(RemarkLines(lf), RemarkLines(crlf));
        }

        [Fact]
        public void 備考欄に収まらない行数を差し込むとレイアウト段階のエラーになる()
        {
            var remarks = string.Join("\n", Enumerable.Range(1, 20).Select(i => $"{i}行目の備考"));

            var ex = Assert.Throws<LayoutComputationException>(() => ComputeInvoice(InvoiceValues(remarks)));

            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Equal(CellAddress.Parse("A29"), ex.CellAddress);
        }

        [Fact]
        public void ConvertToFileは入力ファイルが存在しない場合UtsushiExceptionに変換する()
        {
            // File.OpenReadが投げる生のFileNotFoundExceptionをそのまま漏らすと、
            // ストリーム版(Convert)と異なりUtsushiException階層で一貫して扱えなくなる(要件6.4, 6.5)。
            using var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            using var converter = new ReportPdfConverter(
                new OpenXmlWorkbookReader(),
                new FileSystemReportDefinitionRepository(TestPaths.SampleReportsRoot),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                new FakePdfRenderer(),
                fontResolver);

            var missingPath = Path.Combine(TestPaths.RepositoryRoot, "存在しない.xlsx");
            var outputPath = Path.GetTempFileName();
            try
            {
                var ex = Assert.Throws<InvalidExcelFileException>(() => converter.ConvertToFile(
                    "invoice", missingPath, new Dictionary<string, string>(), outputPath));
                Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            }
            finally
            {
                File.Delete(outputPath);
            }
        }

        /// <summary>SkiaPdfRenderer以外のIPdfRenderer実装を想定したテスト用の最小実装。</summary>
        private sealed class FakePdfRenderer : IPdfRenderer
        {
            public const string Content = "fake-pdf-content";

            public void Render(PagedLayout layout, Stream output)
            {
                var bytes = Encoding.UTF8.GetBytes(Content);
                output.Write(bytes, 0, bytes.Length);
            }
        }

        private static PagedLayout ComputeInvoiceLayout()
        {
            using var converter = CreateDeterministicConverter(out _);
            using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

            return converter.ComputeLayout(
                "invoice",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "株式会社テスト製作所 御中",
                    ["InvoiceNo"] = "INV-2026-0417",
                    ["TotalAmount"] = "¥2,153,800",
                });
        }
    }
}
