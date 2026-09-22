using System;
using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
using Utsushi.Layout.HeaderFooter;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using static Utsushi.Layout.Tests.LayoutFixtures;

namespace Utsushi.Layout.Tests
{
    /// <summary>
    /// ヘッダー/フッターの書式コード展開と配置の検証(要件3.7〜3.9)。
    /// </summary>
    public sealed class HeaderFooterTests
    {
        private static readonly DateTime Timestamp = new(2026, 4, 20, 14, 5, 0);

        private static HeaderFooterContext Context(int pageNumber = 1, int totalPages = 3) =>
            new(pageNumber, totalPages, "請求書", "invoice", Timestamp, FontStyle.Default);

        private static IReadOnlyList<HeaderFooterRun> RunsOf(
            IReadOnlyList<HeaderFooterPart> parts, HeaderFooterSection section) =>
            parts.SingleOrDefault(p => p.Section == section)?.Runs ?? Array.Empty<HeaderFooterRun>();

        private static string TextOf(IReadOnlyList<HeaderFooterPart> parts, HeaderFooterSection section) =>
            string.Concat(RunsOf(parts, section).Select(r => r.Text));

        // -- 書式コードの展開 -------------------------------------------------

        [Fact]
        public void セクション指定で左中央右に振り分ける()
        {
            var parts = HeaderFooterParser.Parse("&L左&C中央&R右", Context());

            Assert.Equal("左", TextOf(parts, HeaderFooterSection.Left));
            Assert.Equal("中央", TextOf(parts, HeaderFooterSection.Center));
            Assert.Equal("右", TextOf(parts, HeaderFooterSection.Right));
        }

        [Fact]
        public void セクション指定が無ければ中央に配置する()
        {
            var parts = HeaderFooterParser.Parse("セクション指定なし", Context());

            var part = Assert.Single(parts);
            Assert.Equal(HeaderFooterSection.Center, part.Section);
        }

        [Fact]
        public void ページ番号と総ページ数を展開する()
        {
            var parts = HeaderFooterParser.Parse("&C&P / &N ページ", Context(pageNumber: 2, totalPages: 5));

            Assert.Equal("2 / 5 ページ", TextOf(parts, HeaderFooterSection.Center));
        }

        [Fact]
        public void 日付と時刻とシート名を展開する()
        {
            var parts = HeaderFooterParser.Parse("&L&D &T&R&A", Context());

            Assert.Equal("2026/04/20 14:05", TextOf(parts, HeaderFooterSection.Left));
            Assert.Equal("請求書", TextOf(parts, HeaderFooterSection.Right));
        }

        [Fact]
        public void 二重のアンパサンドは文字として扱う()
        {
            var parts = HeaderFooterParser.Parse("&CA&&B", Context());

            Assert.Equal("A&B", TextOf(parts, HeaderFooterSection.Center));
        }

        [Fact]
        public void 太字と斜体の切り替えが書式に反映される()
        {
            var parts = HeaderFooterParser.Parse("&C通常&B太字&B通常に戻る", Context());

            var runs = RunsOf(parts, HeaderFooterSection.Center);

            Assert.Equal(3, runs.Count);
            Assert.False(runs[0].Font.Bold);
            Assert.True(runs[1].Font.Bold);
            Assert.False(runs[2].Font.Bold);
            Assert.Equal("太字", runs[1].Text);
        }

        [Fact]
        public void フォント名とサイズの指定が反映される()
        {
            var parts = HeaderFooterParser.Parse("&C&\"MS Mincho,Bold\"&14見出し", Context());

            var run = Assert.Single(RunsOf(parts, HeaderFooterSection.Center));

            Assert.Equal("MS Mincho", run.Font.Name);
            Assert.True(run.Font.Bold);
            Assert.Equal(14.0, run.Font.SizePt);
        }

        [Fact]
        public void 未対応の書式コードは読み飛ばす()
        {
            // &G(画像)は対象外。例外にせず、文字列部分だけを残す。
            var parts = HeaderFooterParser.Parse("&C&Gテキスト", Context());

            Assert.Equal("テキスト", TextOf(parts, HeaderFooterSection.Center));
        }

        [Fact]
        public void 文字色の指定がフォントに反映される()
        {
            var parts = HeaderFooterParser.Parse("&C通常&KFF0000赤字&K0000FF青字", Context());

            var runs = RunsOf(parts, HeaderFooterSection.Center);

            Assert.Equal(3, runs.Count);
            Assert.Equal(FontStyle.Default.Color, runs[0].Font.Color);
            Assert.Equal("通常", runs[0].Text);
            Assert.Equal(new ArgbColor(0xFF, 0xFF, 0x00, 0x00), runs[1].Font.Color);
            Assert.Equal("赤字", runs[1].Text);
            Assert.Equal(new ArgbColor(0xFF, 0x00, 0x00, 0xFF), runs[2].Font.Color);
            Assert.Equal("青字", runs[2].Text);
        }

        [Fact]
        public void 不正な文字色指定は既定色のまま処理を継続する()
        {
            // 16進以外の文字が6文字続く場合、例外にせず既定色のまま後続の文字列を残す。
            var nonHex = HeaderFooterParser.Parse("&C&KZZZZZZ不正", Context());
            var nonHexRun = Assert.Single(RunsOf(nonHex, HeaderFooterSection.Center));
            Assert.Equal(FontStyle.Default.Color, nonHexRun.Font.Color);
            Assert.Equal("不正", nonHexRun.Text);

            // 文字列の残りが6文字に満たないまま終わる場合も、範囲外アクセスにならず
            // 既定色のまま何も出力せずに終了する。
            var truncated = HeaderFooterParser.Parse("&C&KFF", Context());
            Assert.Empty(RunsOf(truncated, HeaderFooterSection.Center));
        }

        [Fact]
        public void 文字色と太字を組み合わせると両方反映される()
        {
            var parts = HeaderFooterParser.Parse("&C&B&KFF0000重要", Context());

            var run = Assert.Single(RunsOf(parts, HeaderFooterSection.Center));

            Assert.True(run.Font.Bold);
            Assert.Equal(new ArgbColor(0xFF, 0xFF, 0x00, 0x00), run.Font.Color);
        }

        [Fact]
        public void セクション切り替えで文字色も既定に戻る()
        {
            var parts = HeaderFooterParser.Parse("&L&KFF0000左は赤&C中央は既定色", Context());

            var leftRun = Assert.Single(RunsOf(parts, HeaderFooterSection.Left));
            var centerRun = Assert.Single(RunsOf(parts, HeaderFooterSection.Center));

            Assert.Equal(new ArgbColor(0xFF, 0xFF, 0x00, 0x00), leftRun.Font.Color);
            Assert.Equal(FontStyle.Default.Color, centerRun.Font.Color);
        }

        [Fact]
        public void 空の指定では何も返さない()
        {
            Assert.Empty(HeaderFooterParser.Parse(null, Context()));
            Assert.Empty(HeaderFooterParser.Parse(string.Empty, Context()));
            Assert.Empty(HeaderFooterParser.Parse("&L&C&R", Context()));
        }

        // -- ページごとの指定の選択(要件3.9) --------------------------------

        [Fact]
        public void 先頭ページのみ別指定が選択される()
        {
            var model = new HeaderFooterModel(
                OddHeader: "通常", OddFooter: null,
                EvenHeader: null, EvenFooter: null,
                FirstHeader: "先頭のみ", FirstFooter: null,
                DifferentOddEven: false, DifferentFirst: true, ScaleWithDocument: true);

            Assert.Equal("先頭のみ", model.GetHeader(1));
            Assert.Equal("通常", model.GetHeader(2));
        }

        [Fact]
        public void 奇数偶数で別指定が選択される()
        {
            var model = new HeaderFooterModel(
                OddHeader: "奇数", OddFooter: null,
                EvenHeader: "偶数", EvenFooter: null,
                FirstHeader: null, FirstFooter: null,
                DifferentOddEven: true, DifferentFirst: false, ScaleWithDocument: true);

            Assert.Equal("奇数", model.GetHeader(1));
            Assert.Equal("偶数", model.GetHeader(2));
            Assert.Equal("奇数", model.GetHeader(3));
        }

        [Fact]
        public void 別指定が無効なら常に通常の指定を使う()
        {
            var model = new HeaderFooterModel(
                OddHeader: "通常", OddFooter: null,
                EvenHeader: "偶数", EvenFooter: null,
                FirstHeader: "先頭", FirstFooter: null,
                DifferentOddEven: false, DifferentFirst: false, ScaleWithDocument: true);

            Assert.Equal("通常", model.GetHeader(1));
            Assert.Equal("通常", model.GetHeader(2));
        }

        // -- レイアウトへの反映 -----------------------------------------------

        [Fact]
        public void ヘッダーは上余白にフッターは下余白に配置される()
        {
            var engine = new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => Timestamp);

            var pageSetup = PageSetupModel.Default with
            {
                HeaderFooter = new HeaderFooterModel(
                    "&Cヘッダー", "&Cフッター", null, null, null, null, false, false, true),
            };

            var sheet = UniformSheet(rows: 2, columns: 2, pageSetup: pageSetup);
            var layout = engine.Compute(ReportModel.Create(Definition(), sheet));

            var page = Assert.Single(layout.Pages);
            var texts = page.Commands.OfType<TextCommand>().ToList();

            var header = texts.Single(t => t.Text == "ヘッダー");
            var footer = texts.Single(t => t.Text == "フッター");
            var body = texts.Single(t => t.Text == "A1");

            // ヘッダーは本文より上、フッターは本文より下
            Assert.True(header.Origin.Y < body.Origin.Y, "ヘッダーは本文より上にあるはず");
            Assert.True(footer.Origin.Y > body.Origin.Y, "フッターは本文より下にあるはず");

            // ヘッダーは上余白の中(本文の開始位置より上)
            Assert.True(
                header.Origin.Y < pageSetup.Margins.TopPt + 20,
                $"ヘッダーは上余白内にあるはず (Y={header.Origin.Y}, 上余白={pageSetup.Margins.TopPt})");

            // フッターは下余白の中
            Assert.True(
                footer.Origin.Y > page.HeightPt - pageSetup.Margins.BottomPt - 20,
                $"フッターは下余白内にあるはず (Y={footer.Origin.Y})");
        }

        [Fact]
        public void 文書と一緒に拡大縮小してもヘッダーフッターの余白位置は変わらない()
        {
            // Excelの「文書と一緒に拡大縮小する」はヘッダー/フッターの文字サイズだけを本文の
            // 縮小率に合わせるものであり、ヘッダー/フッター領域の物理的な余白位置(用紙端からの
            // 距離)自体は動かさない。
            var engine = new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => Timestamp);

            var pageSetup = NoMarginA4(scaling: new PageScaling(50, null, null)) with
            {
                Margins = new PageMargins(0, 0, 60, 60, 40, 40),
                HeaderFooter = new HeaderFooterModel(
                    "&Cヘッダー", "&Cフッター", null, null, null, null, false, false, ScaleWithDocument: true),
            };

            var sheet = UniformSheet(rows: 2, columns: 2, pageSetup: pageSetup);
            var layout = engine.Compute(ReportModel.Create(Definition(), sheet));

            var page = Assert.Single(layout.Pages);
            Assert.Equal(0.5, page.ScaleFactor);

            var header = page.Commands.OfType<TextCommand>().Single(t => t.Text == "ヘッダー");
            var footer = page.Commands.OfType<TextCommand>().Single(t => t.Text == "フッター");

            // ヘッダーのベースラインは、上余白(60pt、未縮小)より下にあるはず。
            // 余白まで0.5倍されるバグがあると、ベースラインが余白の途中(30pt付近)に来てしまう。
            Assert.True(
                header.Origin.Y >= pageSetup.Margins.HeaderPt,
                $"ヘッダーの余白は縮小されないはず (Y={header.Origin.Y}, 余白={pageSetup.Margins.HeaderPt})");

            // フッターのベースラインは、下余白(40pt、未縮小)の外に出てはいけない。
            Assert.True(
                footer.Origin.Y <= page.HeightPt - pageSetup.Margins.FooterPt,
                $"フッターの余白は縮小されないはず (Y={footer.Origin.Y}, 用紙高さ={page.HeightPt}, 余白={pageSetup.Margins.FooterPt})");
        }

        [Fact]
        public void ヘッダーフッターは全ページに出力され総ページ数が展開される()
        {
            var engine = new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => Timestamp);

            var pageSetup = NoMarginA4() with
            {
                Margins = new PageMargins(20, 20, 40, 40, 10, 10),
                HeaderFooter = new HeaderFooterModel(
                    null, "&C&P/&N", null, null, null, null, false, false, true),
            };

            // 3ページに分かれる分量
            var sheet = UniformSheet(rows: 30, columns: 1, rowHeightPt: 100.0, pageSetup: pageSetup);
            var layout = engine.Compute(ReportModel.Create(Definition(), sheet));

            Assert.True(layout.PageCount >= 2);

            foreach (var page in layout.Pages)
            {
                var texts = page.Commands.OfType<TextCommand>().Select(t => t.Text).ToList();
                Assert.Contains($"{page.PageNumber}/{layout.PageCount}", texts);
            }
        }

        [Fact]
        public void ヘッダーフッターが未設定なら描画命令を追加しない()
        {
            var engine = new ReportLayoutEngine(new ApproximateFontMetricsProvider(), () => Timestamp);
            var sheet = UniformSheet(rows: 1, columns: 1, pageSetup: NoMarginA4());

            var page = Assert.Single(engine.Compute(ReportModel.Create(Definition(), sheet)).Pages);

            // 本文の1セルぶんのテキストだけ
            Assert.Single(page.Commands.OfType<TextCommand>());
        }
    }
}
