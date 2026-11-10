using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
    /// サンプル帳票のエンドツーエンド(Parsing → Rendering)ゴールデンテスト(要件8.3、excel2pdf-consolidation 要件6.1)。
    /// <see cref="Excel2Pdf"/> でセル番地を指定して値を書き込む。
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
        /// ヘッダー/フッターの &amp;D(日付)・&amp;T(時刻)を固定するための時刻。
        /// 実時刻のままではゴールデンファイルが毎日変わってしまう。
        /// </summary>
        private static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

        /// <summary>サンプル帳票の列幅の換算に使う最大数字幅。標準フォントからの見積もり(8)と異なるため明示する(design.md 6.1)。</summary>
        private const double SampleMaxDigitWidthPx = 7;

        /// <summary>決定的なフォントメトリクスと固定時刻を使うエンジン(design.md 6.1)。</summary>
        private static Excel2PdfEngine CreateDeterministicEngine(IPdfRenderer? renderer = null)
        {
            // 代替フォント名を省略すると同梱の日本語フォントになるため、CIでも字形欠落の検出(要件5.5)を
            // 有効にしたまま日本語の帳票を描画できる(要件11.1)。
            var fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            return new Excel2PdfEngine(
                new OpenXmlWorkbookReader(),
                new ReportModelBuilder(),
                new CellSubstitutor(),
                new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => FixedTimestamp),
                renderer ?? new SkiaPdfRenderer(new SkiaFontMetricsProvider(fontResolver)),
                fontResolver);
        }

        /// <summary>サンプル帳票を開く。文書名はサンプル名(指定しないとファイル名の <c>template</c> になる)。</summary>
        private static Excel2Pdf Open(string sample, Excel2PdfEngine engine, IReadOnlyDictionary<string, string>? values = null)
        {
            var pdf = new Excel2Pdf(
                TestPaths.SampleTemplate(sample),
                engine,
                new Excel2PdfOptions { MaxDigitWidthPx = SampleMaxDigitWidthPx, DocumentName = sample });
            if (values is not null)
            {
                foreach (var (cell, text) in values)
                {
                    pdf.SetText(cell, text);
                }
            }

            return pdf;
        }

        private static PagedLayout ComputeLayout(string sample, IReadOnlyDictionary<string, string> values)
        {
            using var engine = CreateDeterministicEngine();
            using var pdf = Open(sample, engine, values);
            return pdf.ComputeLayout(checkFit: true);
        }

        public static IEnumerable<object[]> SampleReports()
        {
            yield return new object[]
            {
            "invoice",
            new Dictionary<string, string>
            {
                ["A3"] = "株式会社テスト製作所 御中",
                ["F3"] = "INV-2026-0417",
                ["A5"] = "件名: ゴールデンテスト用案件",
                ["C8"] = "¥2,153,800",
                ["A29"] = "備考:\nゴールデンテストで固定した値です。",
            },
            };

            yield return new object[]
            {
            "receipt",
            new Dictionary<string, string>
            {
                ["A3"] = "株式会社テスト製作所 様",
                ["B5"] = "¥2,153,800",
                ["B7"] = "サンプル案件の代金として",
                ["A18"] = "株式会社テスト製作所 様",
                ["B20"] = "¥2,153,800",
                ["B22"] = "サンプル案件の代金として",
            },
            };

            yield return new object[]
            {
            "delivery-note",
            new Dictionary<string, string>
            {
                ["A3"] = "株式会社テスト製作所 御中",
                ["E3"] = "DN-2026-0088",
                ["E4"] = "2026年4月15日",
            },
            };
        }

        [Theory]
        [MemberData(nameof(SampleReports))]
        public void サンプル帳票のレイアウトが期待結果と一致する(string sample, Dictionary<string, string> values)
        {
            var layout = ComputeLayout(sample, values);

            GoldenFile.Verify(sample, "layout.snapshot.txt", GoldenSnapshot.Create(layout));
        }

        [Theory]
        [MemberData(nameof(SampleReports))]
        public void サンプル帳票からPDFを生成できる(string sample, Dictionary<string, string> values)
        {
            using var engine = CreateDeterministicEngine();
            using var pdf = Open(sample, engine, values);

            var bytes = pdf.ToPdfBytes();

            Assert.True(bytes.Length > 0, "PDFが出力されるはず");
            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        }

        [Fact]
        public void 請求書は1ページに収まる()
        {
            var layout = ComputeLayout("invoice", InvoiceValues());

            Assert.Equal(1, layout.PageCount);
            Assert.Equal("A4", layout.Pages[0].Paper.Name);
        }

        [Fact]
        public void 納品書は複数ページになり印刷タイトルが各ページに繰り返される()
        {
            var layout = ComputeLayout("delivery-note", DeliveryNoteValues());

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
        public void 書き込んだ値が帳票に反映される()
        {
            var values = InvoiceValues();
            values["B12"] = "臨時オプション対応費";

            var texts = ComputeLayout("invoice", values).Pages
                .SelectMany(p => p.Commands.OfType<TextCommand>())
                .Select(t => t.Text)
                .ToList();

            Assert.Contains("株式会社テスト製作所 御中", texts);
            Assert.Contains("INV-2026-0417", texts);
            Assert.Contains("¥2,153,800", texts);
            Assert.Contains("臨時オプション対応費", texts);

            // テンプレートの既定値は残らない
            Assert.DoesNotContain("INV-0000", texts);
            Assert.DoesNotContain("サンプル設計費", texts);
        }

        [Fact]
        public void 結合セルの非アンカーに書き込むとエラーになる()
        {
            // A3:C3(宛先のセル)は結合セルであり、B3・C3はアンカーではない。
            var values = InvoiceValues();
            values["B3"] = "見えなくなる値";

            var ex = Assert.Throws<NonAnchorMergedCellOverrideException>(() => ComputeLayout("invoice", values));

            Assert.Equal(CellAddress.Parse("B3"), ex.CellAddress);
            Assert.Equal(CellAddress.Parse("A3"), ex.AnchorAddress);
            Assert.Equal("invoice", ex.DocumentName);
        }

        [Fact]
        public void 不正なセル番地を指定するとエラーになる()
        {
            using var engine = CreateDeterministicEngine();
            using var pdf = Open("invoice", engine);

            var ex = Assert.Throws<InvalidCellOverrideAddressException>(() => pdf.SetText("not-a-cell", "x"));

            Assert.Equal("not-a-cell", ex.Address);
            Assert.Equal("invoice", ex.DocumentName);
        }

        [Fact]
        public void 複数の印刷範囲はそれぞれ別ページになり範囲外は出力されない()
        {
            // 領収書サンプルは「本紙」と「控え」を別々の印刷範囲として持つ(要件3.6)。
            var layout = ComputeLayout(
                "receipt",
                new Dictionary<string, string>
                {
                    ["A3"] = "株式会社テスト製作所 様",
                    ["B5"] = "¥1,000",
                    ["A18"] = "株式会社テスト製作所 様",
                    ["B20"] = "¥1,000",
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
            var layout = ComputeLayout("delivery-note", DeliveryNoteValues());

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
                Assert.Contains(
                    page.Commands.OfType<TextCommand>(),
                    t => t.Text == "サンプル" && t.Font.Bold);
            }
        }

        [Fact]
        public void 既定以外のIPdfRendererでもファイル出力は一時ファイル経由で確定する()
        {
            // 要件5.4(失敗時に不完全なファイルを残さない)はレンダラーの実装によらず保証されるべき。
            using var engine = CreateDeterministicEngine(new FakePdfRenderer());
            using var pdf = Open("invoice", engine, InvoiceValues());

            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            var outputPath = Path.Combine(directory, "invoice.pdf");

            try
            {
                pdf.Save(outputPath);

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

        [Fact]
        public void 書き込む値の制御文字はSubstitution段階のエラーになる()
        {
            var values = InvoiceValues();
            values["F3"] = "INV\t0001";

            var ex = Assert.Throws<InvalidSubstitutionValueException>(() => ComputeLayout("invoice", values));

            Assert.Equal(ProcessingStage.Substitution, ex.Stage);
            Assert.Equal("invoice", ex.DocumentName);
            Assert.Equal(CellAddress.Parse("F3"), ex.CellAddress);
        }

        [Fact]
        public void 備考欄のCRLF区切りの住所はLF区切りと同じレイアウトになる()
        {
            static IEnumerable<string> RemarkLines(PagedLayout layout) =>
                layout.Pages.SelectMany(p => p.Commands.OfType<TextCommand>())
                    .Where(t => t.Text.Contains("丸の内") || t.Text.Contains("東京都") || t.Text.Contains("担当"))
                    .Select(t => t.Text);

            var lf = ComputeLayout("invoice", InvoiceValues("東京都千代田区丸の内1-1-1\n丸の内ビル10F\n担当: 山田"));
            var crlf = ComputeLayout("invoice", InvoiceValues("東京都千代田区丸の内1-1-1\r\n丸の内ビル10F\r\n担当: 山田"));

            Assert.Equal(new[] { "東京都千代田区丸の内1-1-1", "丸の内ビル10F", "担当: 山田" }, RemarkLines(lf));
            Assert.Equal(RemarkLines(lf), RemarkLines(crlf));
        }

        [Fact]
        public void 備考欄に収まらない行数を書き込むとレイアウト段階のエラーになる()
        {
            var remarks = string.Join("\n", Enumerable.Range(1, 20).Select(i => $"{i}行目の備考"));

            var ex = Assert.Throws<LayoutComputationException>(() => ComputeLayout("invoice", InvoiceValues(remarks)));

            Assert.Equal(ProcessingStage.Layout, ex.Stage);
            Assert.Equal(CellAddress.Parse("A29"), ex.CellAddress);
        }

        [Fact]
        public void 入力ファイルが存在しない場合はUtsushiExceptionに変換する()
        {
            // File.OpenRead が投げる生の FileNotFoundException をそのまま漏らすと、
            // UtsushiException の階層で一貫して扱えなくなる(要件6.4, 6.5)。
            using var engine = CreateDeterministicEngine(new FakePdfRenderer());
            var missingPath = Path.Combine(TestPaths.RepositoryRoot, "存在しない.xlsx");
            using var pdf = new Excel2Pdf(missingPath, engine);
            var outputPath = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N") + ".pdf");

            var ex = Assert.Throws<InvalidExcelFileException>(() => pdf.Save(outputPath));

            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.False(File.Exists(outputPath));
        }

        private static Dictionary<string, string> InvoiceValues(string? remarks = null)
        {
            var values = new Dictionary<string, string>
            {
                ["A3"] = "株式会社テスト製作所 御中",
                ["F3"] = "INV-2026-0417",
                ["C8"] = "¥2,153,800",
            };

            if (remarks is not null)
            {
                values["A29"] = remarks;
            }

            return values;
        }

        private static Dictionary<string, string> DeliveryNoteValues() => new()
        {
            ["A3"] = "株式会社テスト製作所 御中",
            ["E3"] = "DN-2026-0088",
        };

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
    }
}
