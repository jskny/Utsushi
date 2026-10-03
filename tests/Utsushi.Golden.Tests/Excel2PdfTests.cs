using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Utsushi;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Rendering;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// セル番地を指定して値を書き換える簡易API <see cref="Excel2Pdf"/>(要件14)。請求書サンプルを使う。
    /// </summary>
    /// <remarks>
    /// 請求書サンプルのセルの数値書式: C8(ご請求金額)・F12〜F23(金額)は <c>"¥"#,##0</c>、F4・F5(発行日・支払期限)は
    /// <c>yyyy"年"m"月"d"日"</c>。
    /// </remarks>
    public sealed class Excel2PdfTests
    {
        private static readonly string Invoice = TestPaths.SampleTemplate("invoice");

        private static Excel2PdfOptions FallbackFonts() => new() { Fonts = FontResolverOptions.AllowFallback() };

        private static List<string> Texts(Excel2Pdf pdf) =>
            pdf.ComputeLayout(checkFit: false).Pages.SelectMany(p => p.Commands.OfType<TextCommand>()).Select(t => t.Text).ToList();

        private static TextCommand Text(Excel2Pdf pdf, string text) =>
            pdf.ComputeLayout(checkFit: false).Pages.SelectMany(p => p.Commands.OfType<TextCommand>()).Single(t => t.Text == text);

        [Fact]
        public void 文字列と数値と日時を設定してPDFとして保存できる()
        {
            var output = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N") + ".pdf");
            try
            {
                using (var pdf = new Excel2Pdf(Invoice, FallbackFonts()))
                {
                    pdf.SetText("A3", "株式会社テスト製作所 御中");
                    pdf.SetValue("C8", 2153800);
                    pdf.SetValue("F4", new DateTime(2026, 5, 1));
                    pdf.Save(output);
                }

                var bytes = File.ReadAllBytes(output);
                Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
            }
            finally
            {
                File.Delete(output);
            }
        }

        [Fact]
        public void 数値と日時はセルの表示形式で表示する()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.SetValue("C8", 2153800)
               .SetValue("F12", -3000m)
               .SetValue("F4", new DateTime(2026, 5, 1))
               .SetText("F13", "1234");

            var texts = Texts(pdf);

            Assert.Contains("¥2,153,800", texts);   // "¥"#,##0
            Assert.Contains("-¥3,000", texts);
            Assert.Contains("2026年5月1日", texts);  // yyyy"年"m"月"d"日"
            Assert.Contains("1234", texts);          // 文字列には表示形式が効かない
        }

        [Fact]
        public void 数値は右揃えで描き文字列は左揃えで描く()
        {
            // 横位置が「標準」のセルでは、Excel と同じく数値は右、文字列は左に寄る。
            using var number = new Excel2Pdf(Invoice, FallbackFonts());
            number.SetValue("B28", 42);
            using var text = new Excel2Pdf(Invoice, FallbackFonts());
            text.SetText("B28", "42");

            Assert.Equal(TextAnchor.Right, Text(number, "42").Anchor);
            Assert.Equal(TextAnchor.Left, Text(text, "42").Anchor);
        }

        [Fact]
        public void 表記が異なっても同じセルへの設定は最後の値を使う()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.SetText("a3", "最初の値");
            pdf.SetValue("$A$3", 1);
            pdf.SetText("A3", "最後の値");

            var texts = Texts(pdf);

            Assert.Contains("最後の値", texts);
            Assert.DoesNotContain("最初の値", texts);
        }

        [Fact]
        public void 値を変えて何度でも保存できる()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());

            pdf.SetText("A3", "一通目 御中");
            using (var first = new MemoryStream())
            {
                pdf.Save(first);
                Assert.True(first.Length > 0);
            }

            pdf.SetText("A3", "二通目 御中");
            Assert.Contains("二通目 御中", Texts(pdf));
        }

        [Theory]
        [InlineData("A0")]
        [InlineData("1A")]
        [InlineData("")]
        [InlineData("XFE1")]
        public void A1形式でないセル番地は設定の時点でエラーにする(string cell)
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());

            var ex = Assert.Throws<InvalidCellOverrideAddressException>(() => pdf.SetText(cell, "x"));
            Assert.Equal(cell, ex.Address);
            Assert.Throws<InvalidCellOverrideAddressException>(() => pdf.SetValue(cell, 1.0));
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(double.NegativeInfinity)]
        public void 表示できない数値は設定の時点でエラーにする(double value)
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());

            var ex = Assert.Throws<InvalidSubstitutionValueException>(() => pdf.SetValue("C8", value));
            Assert.Equal(CellAddress.Parse("C8"), ex.CellAddress);
        }

        [Fact]
        public void 年1900年より前の日時は設定の時点でエラーにする()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());

            Assert.Throws<InvalidSubstitutionValueException>(() => pdf.SetValue("F4", new DateTime(1899, 12, 31)));
        }

        [Theory]
        [InlineData(1900, 1, 1, 1.0)]
        [InlineData(1900, 2, 28, 59.0)]
        [InlineData(1900, 3, 1, 61.0)]
        [InlineData(2026, 4, 1, 46113.0)]
        public void 日時はExcelの1900年日付システムのシリアル値にする(int year, int month, int day, double expected)
        {
            Assert.Equal(expected, Excel2Pdf.ToExcelSerial(new DateTime(year, month, day)));
        }

        [Fact]
        public void 結合セルの左上以外への設定は保存の時点でエラーにする()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.SetValue("D8", 1000);   // C8:F8 の結合範囲の内側

            Assert.Throws<NonAnchorMergedCellOverrideException>(() => pdf.Save(new MemoryStream()));
        }

        [Fact]
        public void 帳票定義なしでは置換キーを使えない()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());

            Assert.Throws<InvalidOperationException>(() => pdf.SetField("CustomerName", "x"));
        }

        [Fact]
        public void 帳票定義ありでは置換キーとセル番地の両方で設定できる()
        {
            using var pdf = new Excel2Pdf(Invoice, "invoice", TestPaths.SampleReportsRoot, FallbackFonts());
            pdf.SetField("CustomerName", "株式会社テスト製作所 御中")
               .SetField("InvoiceNo", "INV-0001")
               .SetField("TotalAmount", "¥1,000")
               .SetValue("F12", 4321);

            var texts = Texts(pdf);

            Assert.Contains("株式会社テスト製作所 御中", texts);
            Assert.Contains("¥1,000", texts);   // 置換キー(文字列)
            Assert.Contains("¥4,321", texts);   // セル番地の数値(F12 の表示形式 "¥"#,##0)
            Assert.Equal("invoice", pdf.ReportCode);
        }

        [Fact]
        public void 帳票定義ありで必須の置換キーが無ければ保存の時点でエラーにする()
        {
            using var pdf = new Excel2Pdf(Invoice, "invoice", TestPaths.SampleReportsRoot, FallbackFonts());
            pdf.SetField("CustomerName", "株式会社テスト製作所 御中");

            Assert.Throws<RequiredSubstitutionValueMissingException>(() => pdf.Save(new MemoryStream()));
        }

        [Fact]
        public void 収まりの確認は設定した値で行う()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.SetValue("F12", 123456789012345);

            var result = pdf.CheckFit();

            Assert.Contains(result.Issues, i => i.Kind == FitIssueKind.NumberTooWide && i.Cell == CellAddress.Parse("F12"));
        }

        [Fact]
        public void 渡されたコンバータは破棄しない()
        {
            using var converter = ReportPdfConverter.CreateDefault(FontResolverOptions.AllowFallback());

            using (var pdf = new Excel2Pdf(Invoice, converter))
            {
                pdf.SetText("A3", "一通目");
                pdf.Save(new MemoryStream());
            }

            // 同じコンバータで続けて使える。
            using var next = new Excel2Pdf(Invoice, converter);
            next.SetText("A3", "二通目");
            next.Save(new MemoryStream());
        }

        [Fact]
        public void 破棄した後は使えない()
        {
            var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.Dispose();
            pdf.Dispose();

            Assert.Throws<ObjectDisposedException>(() => pdf.SetText("A1", "x"));
            Assert.Throws<ObjectDisposedException>(() => pdf.Save(new MemoryStream()));
        }

        [Fact]
        public void 両立しない設定はコンストラクタでエラーにする()
        {
            Assert.Throws<ArgumentException>(() => new Excel2Pdf(" "));
            Assert.Throws<ArgumentException>(
                () => new Excel2Pdf(Invoice, "invoice", TestPaths.SampleReportsRoot, new Excel2PdfOptions { MaxDigitWidthPx = 8 }));

            using var converter = ReportPdfConverter.CreateDefault(FontResolverOptions.AllowFallback());
            Assert.Throws<ArgumentException>(() => new Excel2Pdf(Invoice, converter, options: FallbackFonts()));
        }

        [Fact]
        public void 帳票定義で折り返しを指定したセルに数値を設定しても行数の検証でエラーにならない()
        {
            // 請求書の Remarks(A29)は overflow: wrap。文字列なら折り返し、行数が高さを超えるとエラー(要件2.14)になるが、
            // 数値は折り返さない(Excel と同じ)。
            using var pdf = new Excel2Pdf(Invoice, "invoice", TestPaths.SampleReportsRoot, FallbackFonts());
            pdf.SetField("CustomerName", "株式会社テスト製作所 御中")
               .SetField("InvoiceNo", "INV-0001")
               .SetField("TotalAmount", "¥1,000")
               .SetValue("A29", 1234.5);

            Assert.Single(Texts(pdf), t => t == "1234.5");
        }

        [Fact]
        public void 時間はシリアル値の日数として設定し負の時間はエラーにする()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.SetValue("B28", TimeSpan.FromHours(36));

            Assert.Contains("1.5", Texts(pdf));   // 表示形式が標準のセル
            Assert.Throws<InvalidSubstitutionValueException>(() => pdf.SetValue("B28", TimeSpan.FromMinutes(-1)));
        }

        [Fact]
        public void 保存のたびに入力ファイルを読み直す()
        {
            var input = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                File.Copy(Invoice, input);
                using var pdf = new Excel2Pdf(input, FallbackFonts());
                Assert.Contains("請求書", Texts(pdf));

                File.Copy(TestPaths.SampleTemplate("receipt"), input, overwrite: true);
                Assert.Contains("領 収 書", Texts(pdf));
            }
            finally
            {
                File.Delete(input);
            }
        }

        [Fact]
        public void 保存に失敗したときは出力ファイルを残さない()
        {
            var output = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N") + ".pdf");
            using var pdf = new Excel2Pdf(Invoice, "invoice", TestPaths.SampleReportsRoot, FallbackFonts());

            Assert.Throws<RequiredSubstitutionValueMissingException>(() => pdf.Save(output));
            Assert.False(File.Exists(output));
        }

        [Fact]
        public void 帳票定義のルートを持つコンバータを渡して帳票定義ありで使える()
        {
            using var converter = ReportPdfConverter.CreateDefault(TestPaths.SampleReportsRoot, FontResolverOptions.AllowFallback());
            using var pdf = new Excel2Pdf(Invoice, converter, "invoice");
            pdf.SetField("CustomerName", "株式会社テスト製作所 御中")
               .SetField("InvoiceNo", "INV-0001")
               .SetField("TotalAmount", "¥1,000");

            pdf.Save(new MemoryStream());

            Assert.Throws<ArgumentException>(() => new Excel2Pdf(Invoice, converter, " "));
        }

        // ---- 要件14.10: バイト列・ストリームの入力、バイト列の出力 ----

        [Fact]
        public void バイト列から作りPDFをバイト列で受け取れる()
        {
            var content = File.ReadAllBytes(Invoice);
            using var pdf = new Excel2Pdf(content, FallbackFonts());

            // 作成の後で元の配列を書き換えても影響しない(作成の時点で複製する)。
            Array.Clear(content, 0, content.Length);

            pdf.SetText("A3", "株式会社テスト製作所 御中");
            var bytes = pdf.ToPdfBytes();

            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
            Assert.Null(pdf.XlsxPath);
            Assert.Equal(string.Empty, pdf.ComputeLayout(checkFit: false).ReportCode);   // 文書名の既定は空文字列(要件12.5)
        }

        [Fact]
        public void ストリームから作るとストリームは読み終えるが閉じない()
        {
            using var stream = File.OpenRead(Invoice);
            using var pdf = new Excel2Pdf(stream, "invoice", TestPaths.SampleReportsRoot, FallbackFonts());

            Assert.True(stream.CanRead);
            Assert.Equal(stream.Length, stream.Position);

            pdf.SetField("CustomerName", "株式会社テスト製作所 御中")
               .SetField("InvoiceNo", "INV-0001")
               .SetField("TotalAmount", "¥1,000");

            // 同じ内容から何度でも変換できる。
            Assert.True(pdf.ToPdfBytes().Length > 0);
            Assert.True(pdf.ToPdfBytes().Length > 0);
        }

        [Fact]
        public void 上限を超える内容のストリームはエラーにする()
        {
            using var stream = new MemoryStream(new byte[1001]);

            var ex = Assert.Throws<InvalidExcelFileException>(() => Excel2Pdf.TemplateSource.FromStream(stream, limit: 1000));
            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
        }

        // ---- 要件14.11: シートの指定 ----

        private static readonly string OrderForm =
            Path.Combine(TestPaths.EdgeCaseSamplesRoot, "edge-order-form", "template.xlsx");

        [Fact]
        public void 帳票定義なしではシート名を指定してアクティブシート以外を変換できる()
        {
            // edge-order-form は「表紙」「注文書」「マスタ(非表示)」の3シートで、アクティブシートは「注文書」。
            using var active = new Excel2Pdf(OrderForm, FallbackFonts());
            Assert.Equal("注文書", active.ComputeLayout(checkFit: false).SheetName);

            using var cover = new Excel2Pdf(OrderForm, new Excel2PdfOptions
            {
                Fonts = FontResolverOptions.AllowFallback(),
                SheetName = "表紙",
            });
            Assert.Equal("表紙", cover.ComputeLayout(checkFit: false).SheetName);
            Assert.Contains("この表紙シートは印刷対象ではない", Texts(cover));
        }

        [Fact]
        public void 存在しないシート名は保存の時点でブック内のシート名を含むエラーにする()
        {
            using var pdf = new Excel2Pdf(OrderForm, new Excel2PdfOptions
            {
                Fonts = FontResolverOptions.AllowFallback(),
                SheetName = "請求書",
            });

            var ex = Assert.Throws<InvalidExcelFileException>(() => pdf.Save(new MemoryStream()));
            Assert.Equal(InvalidExcelFileReason.NoWorksheet, ex.Reason);
            Assert.Contains("注文書", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 帳票定義ありではシート名を指定できない()
        {
            Assert.Throws<ArgumentException>(() => new Excel2Pdf(
                Invoice, "invoice", TestPaths.SampleReportsRoot, new Excel2PdfOptions { SheetName = "請求書" }));
            Assert.Throws<ArgumentException>(() => new Excel2Pdf(Invoice, new Excel2PdfOptions { SheetName = " " }));
        }

        // ---- 要件14.12: 設定の消去 ----

        [Fact]
        public void 設定を消すとテンプレートの値に戻る()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.SetText("A3", "一通目の宛先").SetValue("C8", 1000);

            pdf.Clear("A3");
            var afterOne = Texts(pdf);
            Assert.DoesNotContain("一通目の宛先", afterOne);
            Assert.Contains("株式会社サンプル商事 御中", afterOne);   // テンプレートの値
            Assert.Contains("¥1,000", afterOne);                    // 消していないセルは残る

            pdf.Clear();
            Assert.DoesNotContain("¥1,000", Texts(pdf));

            pdf.Clear("Z99");   // 設定していないセルは何もしない
            Assert.Throws<InvalidCellOverrideAddressException>(() => pdf.Clear("1A"));
        }

        [Fact]
        public void すべての設定を消すと置換キーの設定も消える()
        {
            using var pdf = new Excel2Pdf(Invoice, "invoice", TestPaths.SampleReportsRoot, FallbackFonts());
            pdf.SetField("CustomerName", "株式会社テスト製作所 御中")
               .SetField("InvoiceNo", "INV-0001")
               .SetField("TotalAmount", "¥1,000");

            pdf.Clear();

            Assert.Throws<RequiredSubstitutionValueMissingException>(() => pdf.Save(new MemoryStream()));
        }

        [Fact]
        public void 存在しない入力ファイルは保存の時点でエラーにする()
        {
            using var pdf = new Excel2Pdf(Path.Combine(Path.GetTempPath(), "utsushi-missing-" + Guid.NewGuid().ToString("N") + ".xlsx"), FallbackFonts());

            Assert.Throws<InvalidExcelFileException>(() => pdf.Save(new MemoryStream()));
        }
    }
}
