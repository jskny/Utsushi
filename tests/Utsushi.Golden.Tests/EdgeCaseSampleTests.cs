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
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.Substitution;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// エッジケース検証用の帳票サンプル(<c>samples/edge-cases/</c>)を、差し込み値を変えながら変換する。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 登録済み帳票のゴールデンテスト(<see cref="ReportConversionGoldenTests"/>)が描画命令の完全一致で回帰を検出するのに対し、
    /// こちらは「利用者が普段の Excel 操作で作りがちなテンプレート」と「極端な差し込み値」の組み合わせで、
    /// 変換が成功すること・失敗するべき値が分かりやすい例外で止まること・印刷設定が効いていることを確かめる。
    /// 各サンプルの狙いは <c>samples/edge-cases/README.md</c> を参照。
    /// </para>
    /// <para>
    /// レイアウトには決定的な <see cref="ApproximateFontMetricsProvider"/> を、描画には同梱フォントへのフォールバックを使う。
    /// </para>
    /// </remarks>
    public sealed class EdgeCaseSampleTests
    {
        private const string CoverLetter = "edge-cover-letter";
        private const string SalesList = "edge-sales-list";
        private const string Schedule = "edge-schedule";
        private const string OrderForm = "edge-order-form";

        private static readonly DateTime FixedTimestamp = new(2026, 4, 20, 10, 30, 0, DateTimeKind.Unspecified);

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

        private static string Template(string reportCode) =>
            Path.Combine(TestPaths.EdgeCaseSamplesRoot, reportCode, "template.xlsx");

        private static PagedLayout ComputeLayout(string reportCode, IReadOnlyDictionary<string, string> values)
        {
            using var converter = CreateConverter(TestPaths.EdgeCaseSamplesRoot);
            using var input = File.OpenRead(Template(reportCode));
            return converter.ComputeLayout(reportCode, input, values);
        }

        private static byte[] Convert(string reportCode, IReadOnlyDictionary<string, string> values)
        {
            using var converter = CreateConverter(TestPaths.EdgeCaseSamplesRoot);
            using var input = File.OpenRead(Template(reportCode));
            using var output = new MemoryStream();
            converter.Convert(reportCode, input, values, output);
            return output.ToArray();
        }

        private static List<string> Texts(PageLayout page) =>
            page.Commands.OfType<TextCommand>().Select(t => t.Text).ToList();

        private static List<string> Texts(PagedLayout layout) =>
            layout.Pages.SelectMany(Texts).ToList();

        private static void AssertPdf(byte[] bytes)
        {
            Assert.True(bytes.Length > 0, "PDFが出力されるはず");
            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        }

        // ---- 送付状: 差し込み値の極端な値 ----

        private static Dictionary<string, string> CoverLetterValues(params (string Key, string? Value)[] changes)
        {
            var values = new Dictionary<string, string>
            {
                ["PostalCode"] = "〒100-0001",
                ["Address"] = "東京都千代田区千代田1-1-1\nサンプルビル 10F",
                ["CompanyName"] = "株式会社サンプル商事",
                ["Department"] = "総務部",
                ["PersonName"] = "山田 太郎",
                ["SendDate"] = "2026年4月1日",
                ["DocumentNo"] = "0012",
                ["Sender"] = "株式会社Utsushi\n〒150-0001 東京都渋谷区1-2-3\nTEL 03-0000-0000\n担当: 佐藤",
                ["Item1"] = "ご請求書",
                ["Item1Copies"] = "1",
                ["Item1Note"] = "原本",
                ["Item2"] = "納品書(控)",
                ["Item2Copies"] = "2",
                ["Item2Note"] = "",
                ["AmountGeneral"] = "¥1,234,567",
                ["AmountRight"] = "¥1,234,567",
                ["Remarks"] = "ご不明な点は担当までお問い合わせください。",
                ["LatinFontCell"] = "Calibriのセルに日本語",
                ["UnformattedCell"] = "書式なしのセルに日本語",
                ["ColumnStyleCell"] = "列全体の書式だけのセル",
            };

            foreach (var (key, value) in changes)
            {
                if (value is null)
                {
                    values.Remove(key);
                }
                else
                {
                    values[key] = value;
                }
            }

            return values;
        }

        public static IEnumerable<object[]> AcceptedCoverLetterValues()
        {
            yield return new object[] { "標準的な値", CoverLetterValues() };
            yield return new object[]
            {
                "長い値(縮小・切り取り・はみ出し・折り返しの上限)",
                CoverLetterValues(
                    ("CompanyName", "特定非営利活動法人日本サンプル帳票標準化推進協議会 東日本統括本部 関東第二支部"),
                    ("Department", "経営企画本部 デジタルトランスフォーメーション推進部 第三課"),
                    ("PersonName", "寿限無寿限無五劫の擦り切れ海砂利水魚"),
                    ("Address", "北海道札幌市中央区北一条西二丁目1番地1 サンプルタワー札幌駅前ビルディング 25階 2501号室 気付"),
                    ("Item1", "2026年度上期 業務委託契約書(甲乙双方記名押印済み原本)及び別紙仕様書・見積書・発注書の写し一式"),
                    ("Item1Note", "返送不要・コピーにて保管のこと"),
                    ("AmountGeneral", "¥123,456,789,012"),
                    ("Remarks", "1行目\n2行目\n3行目\n4行目\n5行目"),
                    ("Sender", "株式会社Utsushi\n〒150-0001\n東京都渋谷区神宮前1-2-3\nTEL 03-0000-0000")),
            };
            yield return new object[]
            {
                "必須の値だけ",
                new Dictionary<string, string>
                {
                    ["PostalCode"] = "〒100-0001", ["Address"] = "東京都千代田区", ["CompanyName"] = "株式会社A",
                    ["SendDate"] = "2026年4月1日", ["DocumentNo"] = "1", ["Sender"] = "株式会社Utsushi",
                    ["Item1"] = "見積書", ["Item1Copies"] = "1",
                },
            };
            yield return new object[]
            {
                "任意の値が空文字・空白のみ",
                CoverLetterValues(("PersonName", ""), ("Department", "   "), ("Remarks", "")),
            };
            yield return new object[]
            {
                "折り返さないセルの改行",
                CoverLetterValues(("PersonName", "山田\r\n太郎"), ("Address", "東京都\r\n千代田区")),
            };
            yield return new object[]
            {
                "空白を含まない長いURL",
                CoverLetterValues(("Remarks", "https://example.com/" + new string('a', 150) + "?query=1")),
            };
            yield return new object[]
            {
                "ノーブレークスペース・全角空白",
                CoverLetterValues(("PersonName", "山田 太郎　様")),
            };
            yield return new object[]
            {
                "半角カナ・丸数字・ローマ数字・記号・JIS第1〜2水準外の漢字(髙・﨑)",
                CoverLetterValues(
                    ("CompanyName", "髙島屋 﨑山商店"),
                    ("Department", "ｶﾌﾞｼｷｶﾞｲｼｬ ㈱ ①②③ Ⅳ"),
                    ("Address", "東京都〜～ 1−2−3 ‐ー―\n♯♭♪ ℃ № ㎡")),
            };
            yield return new object[]
            {
                "マイナスの金額",
                CoverLetterValues(("AmountRight", "-¥1,000"), ("AmountGeneral", "▲1,000")),
            };
        }

        [Theory]
        [MemberData(nameof(AcceptedCoverLetterValues))]
        public void 送付状は極端な差し込み値でも例外にならず1ページのPDFになる(string scenario, Dictionary<string, string> values)
        {
            var layout = ComputeLayout(CoverLetter, values);
            Assert.True(layout.PageCount == 1, $"{scenario}: 1ページに収まるはず(実際 {layout.PageCount} ページ)");

            AssertPdf(Convert(CoverLetter, values));
        }

        [Fact]
        public void 折り返さないセルの改行は表示せず1行につなげる()
        {
            var layout = ComputeLayout(CoverLetter, CoverLetterValues(("PersonName", "山田\r\n太郎")));

            Assert.Contains("山田太郎", Texts(layout));
        }

        [Fact]
        public void 空白を含まない長いURLは文字単位で折り返す()
        {
            var url = "https://example.com/" + new string('a', 150) + "?query=1";

            var texts = Texts(ComputeLayout(CoverLetter, CoverLetterValues(("Remarks", url))));

            // 備考の各行は連続した描画命令になる。先頭行から、つなげると URL 全体になるまでを集める。
            var first = texts.FindIndex(t => t.StartsWith("https://example.com/", StringComparison.Ordinal));
            Assert.True(first >= 0, "備考の先頭行が描画されるはず");
            var lines = new List<string>();
            for (var i = first; i < texts.Count && string.Concat(lines).Length < url.Length; i++)
            {
                lines.Add(texts[i]);
            }

            Assert.True(lines.Count >= 2, $"複数行に折り返されるはず(実際 {lines.Count} 行)");
            Assert.Equal(url, string.Concat(lines));
        }

        [Fact]
        public void 送付状の住所が折り返しでセルの高さを超えると例外で止まる()
        {
            var values = CoverLetterValues(("Address", "1\n2\n3\n4\n5\n6"));

            var ex = Assert.Throws<LayoutComputationException>(() => ComputeLayout(CoverLetter, values));

            Assert.Equal(CellAddress.Parse("A4"), ex.CellAddress);
        }

        public static IEnumerable<object[]> RejectedSubstitutionValues()
        {
            yield return new object[] { "タブ文字(表計算ソフトからのコピー)", "Address", "東京都千代田区\t千代田1-1-1" };
            yield return new object[] { "ゼロ幅スペース", "PersonName", "山田\u200B太郎" };
            yield return new object[] { "BOM", "PersonName", "\uFEFF山田太郎" };

            // 対になっていないサロゲートは xUnit がテストケースをシリアライズする際に U+FFFD へ化ける
            // (Visual Studio のテストエクスプローラーで実行した場合)ため、別の [Fact] で確かめる。
        }

        [Theory]
        [MemberData(nameof(RejectedSubstitutionValues))]
        public void 描画できない文字を含む差し込み値は例外で止まる(string scenario, string key, string value)
        {
            var values = CoverLetterValues((key, value));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(() => ComputeLayout(CoverLetter, values));

            Assert.True(ex.Target == key, $"{scenario}: 例外の Target は置換キー {key} のはず(実際 {ex.Target})");
        }

        [Fact]
        public void 対になっていないサロゲートを含む差し込み値は例外で止まる()
        {
            var values = CoverLetterValues(("PersonName", "山田\uD842"));

            var ex = Assert.Throws<InvalidSubstitutionValueException>(() => ComputeLayout(CoverLetter, values));

            Assert.Equal("PersonName", ex.Target);
        }

        [Fact]
        public void 必須の値が空文字なら例外で止まる()
        {
            var ex = Assert.Throws<RequiredSubstitutionValueMissingException>(
                () => ComputeLayout(CoverLetter, CoverLetterValues(("CompanyName", ""))));

            Assert.Equal(CellAddress.Parse("A7"), ex.CellAddress);
        }

        [Fact]
        public void 必須の値が渡されなければ例外で止まる()
        {
            var ex = Assert.Throws<RequiredSubstitutionValueMissingException>(
                () => ComputeLayout(CoverLetter, CoverLetterValues(("Address", null))));

            Assert.Equal(CellAddress.Parse("A4"), ex.CellAddress);
        }

        [Fact]
        public void 帳票定義に無い置換キーは例外で止まる()
        {
            var ex = Assert.Throws<SubstitutionKeyNotFoundException>(
                () => ComputeLayout(CoverLetter, CoverLetterValues(("CustomerNmae", "綴りを誤ったキー"))));

            Assert.Equal("CustomerNmae", ex.Key);
        }

        [Fact]
        public void どのフォントにも無い絵文字は豆腐にせず例外で止まる()
        {
            // 外字用の代替フォントの既定(同梱のBIZ UDPゴシック → IPAmj明朝)に絵文字が無いことに依存する。
            // 既定の一覧に絵文字を持つフォントを加えた場合は、このテストを見直す。
            var values = CoverLetterValues(("PersonName", "山田 太郎 \U0001F647"));

            Assert.Throws<MissingGlyphException>(() => Convert(CoverLetter, values));
        }

        // ---- 売上一覧: 「すべての列を1ページに印刷」・非表示の行と列・タイトル行 ----

        private static PagedLayout ComputeSalesListLayout() =>
            ComputeLayout(SalesList, new Dictionary<string, string>
            {
                ["Period"] = "2026年4月1日〜2026年4月30日",
                ["Department"] = "東日本営業統括本部 首都圏第二営業部 法人営業課",
            });

        [Fact]
        public void 売上一覧はすべての列を1ページの幅に縮小し縦方向にだけ改ページする()
        {
            var layout = ComputeSalesListLayout();

            Assert.Equal(6, layout.PageCount);
            Assert.All(layout.Pages, page =>
            {
                Assert.Equal((1, 11), page.ColumnRange);
                Assert.Equal(PageOrientation.Landscape, page.Orientation);
                Assert.InRange(page.ScaleFactor, 0.10, 0.99);
            });

            // 「次のページ数に合わせて印刷」のときは手動の改ページ(54行目の後)を無視する(Excel と同じ)
            Assert.DoesNotContain(layout.Pages, page => page.RowRange.Last == 54);
        }

        [Fact]
        public void 売上一覧の見出し行は全ページに繰り返し非表示の列と行は印刷しない()
        {
            var layout = ComputeSalesListLayout();

            Assert.All(layout.Pages, page => Assert.Contains("得意先", Texts(page)));

            var texts = Texts(layout);
            Assert.DoesNotContain("原価(社外秘)", texts);

            // 25行ごとに隠した行(No.25, 50, …)と、高さ0の行(No.11)は出ない。
            // No. 以外の列は数量が1〜9、金額が「¥」付き、日付が「/」区切りのため、No. の値と文字列が一致しない。
            foreach (var hiddenNo in new[] { "11", "25", "50", "75", "100", "125", "150" })
            {
                Assert.DoesNotContain(hiddenNo, texts);
            }

            Assert.Contains("24", texts);
            Assert.Contains("26", texts);

            // 数式の計算結果(Excel が保存したキャッシュ値)が出る
            Assert.Contains("¥41,115,780", texts);
        }

        // ---- 工程表: 行と列の両方向の改ページ・印刷タイトル列・ページの方向・先頭ページ番号 ----

        [Fact]
        public void 工程表は列方向を先に改ページしタイトル列を繰り返す()
        {
            var layout = ComputeLayout(Schedule, new Dictionary<string, string>
            {
                ["ProjectName"] = "サンプル新社屋建設工事",
                ["Author"] = "作成: 工務部 佐藤",
            });

            Assert.Equal(8, layout.PageCount);

            // 「左から右」(overThenDown): 同じ行範囲の列ページが続く。列の手動改ページは24列目(X)の後
            Assert.Equal((1, 24), layout.Pages[0].ColumnRange);
            Assert.Equal((25, 43), layout.Pages[1].ColumnRange);
            Assert.Equal(layout.Pages[0].RowRange, layout.Pages[1].RowRange);

            // 行の手動改ページ(29行目の後)
            Assert.Contains(layout.Pages, page => page.RowRange.Last == 29);

            // タイトル列(A:C)は右側のページにも出る
            Assert.Contains("作業項目 001", Texts(layout.Pages[1]));

            // 先頭ページ番号5から数え、先頭ページだけ別のヘッダー/フッターを使う
            Assert.Contains("工程表(表紙ページ)", Texts(layout.Pages[0]));
            Assert.Contains("- 5 -", Texts(layout.Pages[0]));
            // 総ページ数(&N)を最後のページの番号にすること(要件3.15)は Excel 実機との突き合わせが未実施。
            Assert.Contains("6 / 12", Texts(layout.Pages[1]));
        }

        // ---- 注文書: よく使われる Excel の機能 ----

        private static readonly IReadOnlyDictionary<string, string> OrderFormValues = new Dictionary<string, string>
        {
            ["SupplierName"] = "株式会社サンプル文具オフィスサプライ東日本ロジスティクスセンター",
            ["OrderNo"] = "PO-2026-0001",
            ["PaymentTerms"] = "月末締め翌月末払い(プルダウンの選択肢に無い値)",
            ["Approver"] = "部長 山田",
        };

        [Fact]
        public void 注文書は数式の計算結果と各種の表示形式を出力する()
        {
            var layout = ComputeLayout(OrderForm, OrderFormValues);

            Assert.Equal(1, layout.PageCount);

            var texts = Texts(layout);
            Assert.Contains("¥77,660", texts);        // =G18+G19 のキャッシュ値
            Assert.Contains("#DIV/0!", texts);        // エラー値
            Assert.Contains("令和8年4月1日", texts);  // [$-ja-JP]ggge"年"m"月"d"日"
            Assert.Contains("R8.4.1", texts);         // ge.m.d
            Assert.Contains("36:00", texts);          // [h]:mm
            Assert.Contains("1,235千円", texts);      // #,##0,"千円"
            Assert.Contains("▲5", texts);             // 0;"▲"0
            Assert.Contains("0012", texts);           // 0000
            Assert.Contains("2026年4月30日 (厳守)", texts); // リッチテキストは1つの書式の文字列になる(ガイドに記載の既知の制限)
        }

        [Fact]
        public void 注文書はunsupportedElementsがignoreなら画像だけを出力しグラフで止まらない()
        {
            var layout = ComputeLayout(OrderForm, OrderFormValues);

            var images = layout.Pages.SelectMany(p => p.Commands.OfType<ImageCommand>()).ToList();
            Assert.Equal(2, images.Count);
            Assert.Contains(images, i => i.ContentType == "image/png");
            Assert.Contains(images, i => i.ContentType == "image/jpeg");

            AssertPdf(Convert(OrderForm, OrderFormValues));
        }

        [Fact]
        public void 注文書をunsupportedElementsがerrorの帳票定義で変換するとグラフで止まる()
        {
            var root = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            var directory = Path.Combine(root, OrderForm);
            Directory.CreateDirectory(directory);
            try
            {
                File.Copy(Template(OrderForm), Path.Combine(directory, "template.xlsx"));
                var definition = File.ReadAllText(Path.Combine(TestPaths.EdgeCaseSamplesRoot, OrderForm, "definition.json"))
                    .Replace("\"unsupportedElements\": \"ignore\"", "\"unsupportedElements\": \"error\"", StringComparison.Ordinal);
                Assert.Contains("\"error\"", definition, StringComparison.Ordinal);
                File.WriteAllText(Path.Combine(directory, "definition.json"), definition);

                using var converter = CreateConverter(root);
                using var input = File.OpenRead(Path.Combine(directory, "template.xlsx"));

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(
                    () => converter.ComputeLayout(OrderForm, input, OrderFormValues));

                Assert.Equal("Chart", ex.ElementKind);
            }
            finally
            {
                Directory.Delete(root, recursive: true);
            }
        }

        // ---- 帳票定義なしの変換 ----

        [Theory]
        [InlineData(CoverLetter, 1)]
        [InlineData(SalesList, 6)]
        [InlineData(Schedule, 8)]
        [InlineData(OrderForm, 1)]
        public void エッジケースのテンプレートは帳票定義なしでも変換できる(string reportCode, int expectedPages)
        {
            using var converter = CreateConverter(TestPaths.EdgeCaseSamplesRoot);

            using (var input = File.OpenRead(Template(reportCode)))
            {
                var layout = converter.ComputeLayoutWithoutDefinition(input);
                Assert.Equal(expectedPages, layout.PageCount);
            }

            using (var input = File.OpenRead(Template(reportCode)))
            using (var output = new MemoryStream())
            {
                converter.ConvertWithoutDefinition(input, output);
                AssertPdf(output.ToArray());
            }
        }

        [Fact]
        public void 帳票定義なしでは保存時に開いていたシートを変換し非表示のシートは使わない()
        {
            using var converter = CreateConverter(TestPaths.EdgeCaseSamplesRoot);
            using var input = File.OpenRead(Template(OrderForm));

            var layout = converter.ComputeLayoutWithoutDefinition(input);

            Assert.Equal("注文書", layout.SheetName);
        }

        // ---- エッジケース検証レポートの検出事項への対応(要件1.10, 1.11, 4.12, 13) ----

        [Fact]
        public void 負の数を赤で表示する表示形式の値は赤で描く()
        {
            var texts = ComputeLayout(OrderForm, OrderFormValues).Pages.SelectMany(p => p.Commands.OfType<TextCommand>()).ToList();

            var red = new ArgbColor(0xFF, 0xFF, 0x00, 0x00);
            Assert.All(texts.Where(t => t.Text.StartsWith("-3,000", StringComparison.Ordinal)), t => Assert.Equal(red, t.Font.Color));
            Assert.Equal(2, texts.Count(t => t.Text.StartsWith("-3,000", StringComparison.Ordinal)));
            Assert.NotEqual(red, texts.First(t => t.Text.StartsWith("4,800", StringComparison.Ordinal)).Font.Color);
        }

        [Fact]
        public void テンプレートに無いセルと列全体の書式だけのセルへの差し込みは位置の書式で描く()
        {
            var texts = ComputeLayout(CoverLetter, CoverLetterValues()).Pages.SelectMany(p => p.Commands.OfType<TextCommand>()).ToList();

            // B32 はファイルにセルが無い → ブックの標準の書式(游ゴシック)。以前は Calibri だった。
            Assert.Equal("游ゴシック", texts.Single(t => t.Text == "書式なしのセルに日本語").Font.Name);

            // C33 は列全体に太字・赤を設定した列のセル。
            var columnStyled = texts.Single(t => t.Text == "列全体の書式だけのセル").Font;
            Assert.True(columnStyled.Bold);
            Assert.Equal(new ArgbColor(0xFF, 0xC0, 0x00, 0x00), columnStyled.Color);
        }

        [Fact]
        public void 行全体の塗りつぶしはセルの無い位置にも描く()
        {
            var page = Assert.Single(ComputeLayout(CoverLetter, CoverLetterValues()).Pages);

            // 2行目(表題の下の帯)は A〜H 列の8セルぶん塗りつぶされる。
            var bandFills = page.Commands.OfType<FillRectCommand>()
                .GroupBy(f => (Math.Round(f.Rect.Top, 2), Math.Round(f.Rect.Height, 2)))
                .Where(g => Math.Abs(g.Key.Item2 - 6.0) < 0.01)
                .ToList();
            Assert.Equal(8, Assert.Single(bandFills).Count());
        }

        [Fact]
        public void 収まりの確認で隣の値と重なる文字と収まらない数値を検出しPDFは変えない()
        {
            using var converter = CreateConverter(TestPaths.EdgeCaseSamplesRoot);

            IReadOnlyList<FitIssue> issues;
            using (var input = File.OpenRead(Template(OrderForm)))
            {
                issues = converter.CheckFit(OrderForm, input, OrderFormValues);
            }

            // 狭い A 列の見出し「支払条件」が、右隣の B5「月末締め…」に重なる(検証レポート5-1)。
            var label = Assert.Single(issues, i => i.Cell == CellAddress.Parse("A5"));
            Assert.Equal(FitIssueKind.OverlapsNeighborValue, label.Kind);
            Assert.False(label.IsSubstituted);
            Assert.Contains("B5", label.Message, StringComparison.Ordinal);

            // 差し込んだ値(B5、置換キー PaymentTerms)は結合範囲 B5:C5 の幅で切れる。
            var payment = Assert.Single(issues, i => i.Cell == CellAddress.Parse("B5"));
            Assert.Equal(FitIssueKind.Clipped, payment.Kind);
            Assert.Equal("PaymentTerms", payment.SubstitutionKey);

            // ComputeLayout の結果と同じ(CheckFit は ComputeLayout の FitIssues を返すだけで、描画は変えない)。
            using (var input = File.OpenRead(Template(OrderForm)))
            {
                var layout = converter.ComputeLayout(OrderForm, input, OrderFormValues);
                Assert.Equal(issues, layout.FitIssues);
            }
        }

        [Fact]
        public void 収まりの確認は売上一覧の長い得意先名の重なりを検出する()
        {
            var issues = ComputeSalesListLayout().FitIssues;

            Assert.Contains(issues, i => i.Kind == FitIssueKind.OverlapsNeighborValue && i.Text == "合同会社見本ソリューションズ東日本支社");
        }

        [Fact]
        public void 収まりの確認は帳票定義なしでも使え変換がエラーになる入力では同じ例外を送出する()
        {
            using var converter = CreateConverter(TestPaths.EdgeCaseSamplesRoot);

            using (var input = File.OpenRead(Template(OrderForm)))
            {
                Assert.Contains(converter.CheckFitWithoutDefinition(input), i => i.Cell == CellAddress.Parse("A5"));
            }

            using (var input = File.OpenRead(Template(CoverLetter)))
            {
                Assert.Throws<LayoutComputationException>(
                    () => converter.CheckFit(CoverLetter, input, CoverLetterValues(("Address", "1\n2\n3\n4\n5\n6"))));
            }
        }
    }
}
