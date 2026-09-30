using System;
using System.IO;
using DocumentFormat.OpenXml;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// セル・ヘッダー/フッター・書式の中身の大きさに対する安全弁の検証。
    /// <list type="bullet">
    /// <item>要件6.8: セルの文字列(共有文字列・インライン文字列)の長さ <c>MaxCellTextLength</c>(32,767文字)</item>
    /// <item>要件6.8: ヘッダー/フッター1つの長さ <c>MaxHeaderFooterTextLength</c>(1,024文字)</item>
    /// <item>要件6.8: 数値書式コードの長さ <c>StyleTable.MaxNumberFormatCodeLength</c>(255文字。超過は既定書式にフォールバック)</item>
    /// <item>要件6.9: 拡大縮小率 <c>pageSetup/@scale</c> を 10〜400% に丸める(0 は 100%)</item>
    /// </list>
    /// いずれも上限ちょうどは通り、1超えると失敗(またはフォールバック)することを確認する。
    /// </summary>
    public sealed class ContentLimitTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        // --- 文字列の長さ(ヘルパー) -------------------------------------------------

        [Theory]
        [InlineData(0, 0)]
        [InlineData(5, 5)]
        [InlineData(4, 5)]
        public void 文字列の長さが上限以下なら例外にならない(int length, int maxLength)
        {
            OpenXmlWorkbookReader.EnsureTextLengthWithinLimit(new string('あ', length), maxLength, "テスト文字列", "report");
        }

        [Theory]
        [InlineData(1, 0)]
        [InlineData(6, 5)]
        public void 文字列の長さが上限を1超えるとTooLargeになる(int length, int maxLength)
        {
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureTextLengthWithinLimit(new string('あ', length), maxLength, "テスト文字列", "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Contains("テスト文字列", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void 文字列の長さの上限はExcel自体の上限に合わせてある()
        {
            Assert.Equal(32767, OpenXmlWorkbookReader.MaxCellTextLength);
            Assert.Equal(1024, OpenXmlWorkbookReader.MaxHeaderFooterTextLength);
            Assert.Equal(255, StyleTable.MaxNumberFormatCodeLength);
        }

        // --- インライン文字列(結合テスト) -------------------------------------------

        [Fact]
        public void インライン文字列が上限ちょうどの長さなら読み取れる()
        {
            var text = new string('字', OpenXmlWorkbookReader.MaxCellTextLength);
            var path = SafetyLimitWorkbookFixtures.CreateWithInlineText(text);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(text, sheet.GetCell(CellAddress.Parse("A1"))?.DisplayValue);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void インライン文字列が上限を1文字超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithInlineText(new string('字', OpenXmlWorkbookReader.MaxCellTextLength + 1));
            try
            {
                AssertTooLarge(path, SafetyLimitWorkbookFixtures.SheetName);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 共有文字列(結合テスト) -------------------------------------------------

        [Fact]
        public void 共有文字列が上限ちょうどの長さなら読み取れる()
        {
            var text = new string('字', OpenXmlWorkbookReader.MaxCellTextLength);
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(new X.SharedStringItem(new X.Text(text)));
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(text, sheet.GetCell(CellAddress.Parse("A2"))?.DisplayValue);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 共有文字列が上限を1文字超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(
                new X.SharedStringItem(new X.Text(new string('字', OpenXmlWorkbookReader.MaxCellTextLength + 1))));
            try
            {
                AssertTooLarge(path, "共有文字列");
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void セルから参照されない共有文字列も長さを検査する()
        {
            // 2つ目(どのセルからも参照しない)だけが長すぎる。共有文字列は全件を先に読むため、参照の有無によらず落ちる。
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(
                new X.SharedStringItem(new X.Text("短い")),
                new X.SharedStringItem(new X.Text(new string('字', OpenXmlWorkbookReader.MaxCellTextLength + 1))));
            try
            {
                AssertTooLarge(path, "共有文字列");
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void リッチテキストの共有文字列は連結後の長さが上限ちょうどなら読み取れる()
        {
            var half = OpenXmlWorkbookReader.MaxCellTextLength / 2;
            var first = new string('前', half);
            var second = new string('後', OpenXmlWorkbookReader.MaxCellTextLength - half);
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(
                new X.SharedStringItem(new X.Run(new X.Text(first)), new X.Run(new X.Text(second))));
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(first + second, sheet.GetCell(CellAddress.Parse("A2"))?.DisplayValue);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void リッチテキストの共有文字列は個々のrunが上限内でも連結後に上限を1文字超えるとTooLargeになる()
        {
            var half = OpenXmlWorkbookReader.MaxCellTextLength / 2;
            var path = SafetyLimitWorkbookFixtures.CreateWithSharedStrings(
                new X.SharedStringItem(
                    new X.Run(new X.Text(new string('前', half))),
                    new X.Run(new X.Text(new string('後', OpenXmlWorkbookReader.MaxCellTextLength - half + 1)))));
            try
            {
                AssertTooLarge(path, "共有文字列");
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- ヘッダー/フッター(結合テスト) -------------------------------------------

        /// <summary>ヘッダー/フッターの6種類(奇数・偶数・先頭ページ × ヘッダー・フッター)。</summary>
        public static TheoryData<string> HeaderFooterKinds() => new()
        {
            "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter",
        };

        [Theory]
        [MemberData(nameof(HeaderFooterKinds))]
        public void ヘッダーフッターが上限ちょうどの長さなら読み取れる(string kind)
        {
            var text = "&C" + new string('字', OpenXmlWorkbookReader.MaxHeaderFooterTextLength - 2);
            var path = SafetyLimitWorkbookFixtures.CreateWithHeaderFooter(HeaderFooterElement(kind, text));
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(text, HeaderFooterText(sheet.PageSetup.HeaderFooter, kind));
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [MemberData(nameof(HeaderFooterKinds))]
        public void ヘッダーフッターが上限を1文字超えるとTooLargeになる(string kind)
        {
            var text = "&C" + new string('字', OpenXmlWorkbookReader.MaxHeaderFooterTextLength - 1);
            var path = SafetyLimitWorkbookFixtures.CreateWithHeaderFooter(HeaderFooterElement(kind, text));
            try
            {
                AssertTooLarge(path, "ヘッダー/フッター");
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 拡大縮小率(結合テスト) -------------------------------------------------

        [Theory]
        [InlineData(100U, 100)]
        [InlineData(10U, 10)] // 下限ちょうど
        [InlineData(9U, 10)] // 下限-1 は下限に丸める
        [InlineData(1U, 10)]
        [InlineData(0U, 100)] // 0 は従来どおり等倍として扱う(丸めより前に判定する)
        [InlineData(400U, 400)] // 上限ちょうど
        [InlineData(401U, 400)] // 上限+1 は上限に丸める
        [InlineData(999U, 400)]
        [InlineData(1000U, 400)]
        [InlineData(1001U, 400)] // 1000 を超える値も上限に丸める
        [InlineData(2147483647U, 400)]
        [InlineData(2147483648U, 400)] // int に収まらない値でも負にならない
        [InlineData(4294967295U, 400)]
        public void 拡大縮小率は10から400の範囲に丸める(uint scale, int expectedPercent)
        {
            var path = SafetyLimitWorkbookFixtures.CreateWithScale(scale);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(expectedPercent, sheet.PageSetup.Scaling.ScalePercent);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 拡大縮小率の範囲はExcelと同じ10から400()
        {
            Assert.Equal(10, OpenXmlWorkbookReader.MinPrintScalePercent);
            Assert.Equal(400, OpenXmlWorkbookReader.MaxPrintScalePercent);
        }

        // --- 数値書式コードの長さ(結合テスト。例外にせずフォールバックする) -----------

        [Fact]
        public void 数値書式コードが上限ちょうどの長さなら書式として使われる()
        {
            var code = LiteralSuffixFormat(StyleTable.MaxNumberFormatCodeLength);
            var path = SafetyLimitWorkbookFixtures.CreateWithNumberFormat(code);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var cell = sheet.GetCell(CellAddress.Parse("A2"));

                Assert.NotNull(cell);
                Assert.Equal(code, cell!.Style.NumberFormat);
                Assert.NotEqual("1234.5", cell.DisplayValue);
                Assert.StartsWith("1235", cell.DisplayValue, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 数値書式コードが上限を1文字超えると例外にせず既定書式にフォールバックする()
        {
            var code = LiteralSuffixFormat(StyleTable.MaxNumberFormatCodeLength + 1);
            var path = SafetyLimitWorkbookFixtures.CreateWithNumberFormat(code);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var cell = sheet.GetCell(CellAddress.Parse("A2"));

                Assert.NotNull(cell);
                Assert.Null(cell!.Style.NumberFormat);
                Assert.Equal("1234.5", cell.DisplayValue);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- ヘルパー ----------------------------------------------------------------

        private void AssertTooLarge(string path, string expectedInMessage)
        {
            using var stream = File.OpenRead(path);
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "report")));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            Assert.Equal("report", ex.ReportCode);
            Assert.Contains(expectedInMessage, ex.Message, StringComparison.Ordinal);
        }

        /// <summary>整数表示の後ろに引用符付きのリテラル文字列を続けた、ちょうど <paramref name="length"/> 文字の書式コード。</summary>
        private static string LiteralSuffixFormat(int length)
        {
            // 0"xxx…x" = 1 + 1 + (length - 3) + 1 文字。
            var code = "0\"" + new string('x', length - 3) + "\"";
            Assert.Equal(length, code.Length);
            return code;
        }

        private static OpenXmlLeafTextElement HeaderFooterElement(string kind, string text) => kind switch
        {
            "oddHeader" => new X.OddHeader(text),
            "oddFooter" => new X.OddFooter(text),
            "evenHeader" => new X.EvenHeader(text),
            "evenFooter" => new X.EvenFooter(text),
            "firstHeader" => new X.FirstHeader(text),
            "firstFooter" => new X.FirstFooter(text),
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null),
        };

        private static string? HeaderFooterText(HeaderFooterModel model, string kind) => kind switch
        {
            "oddHeader" => model.OddHeader,
            "oddFooter" => model.OddFooter,
            "evenHeader" => model.EvenHeader,
            "evenFooter" => model.EvenFooter,
            "firstHeader" => model.FirstHeader,
            "firstFooter" => model.FirstFooter,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null),
        };
    }
}
