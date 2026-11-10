using System;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// ファイル由来・呼び出し元由来の文字列の扱いの安全弁(excel2pdf-consolidation タスク8.9 の security-reviewer指摘)。
    /// </summary>
    public sealed class InputHardeningTests
    {
        [Fact]
        public void 有効な範囲より十分長いセル範囲とセル番地の文字列は分割せずに拒否する()
        {
            var longRange = string.Join(":", new string[10_000]);

            Assert.False(CellRange.TryParse(longRange, out _));
            Assert.False(CellAddress.TryParse(new string('A', 100) + "1", out _));
            Assert.True(CellRange.TryParse("$XFD$1048575:$XFD$1048576", out _));
        }

        [Fact]
        public void 印刷タイトルの長すぎる参照は読み飛ばす()
        {
            var titles = DefinedNameParser.ParsePrintTitles("Sheet1!" + string.Join(":", new string[10_000]) + ",Sheet1!$1:$2");

            Assert.Equal(1, titles.FirstRow);
            Assert.Equal(2, titles.LastRow);
        }

        [Fact]
        public void 診断にしか使わない値の範囲外の数値は既定値にする()
        {
            Assert.Equal(0U, LenientValue.Read<uint>(() => throw new OverflowException(), 0U));
            Assert.Equal(0U, LenientValue.Read<uint>(() => throw new FormatException(), 0U));
        }

        [Theory]
        [InlineData("請求書", true)]
        [InlineData("1234567890123456789012345678901", true)]
        [InlineData("12345678901234567890123456789012", false)]
        [InlineData("a:b", false)]
        [InlineData("a\nb", false)]
        [InlineData("a\u2028b", false)]
        [InlineData("", false)]
        public void シート名はExcelの規則で確かめる(string name, bool expected)
        {
            Assert.Equal(expected, DiagnosticText.IsValidSheetName(name));
        }

        [Fact]
        public void 診断用の文字列は文書名とシート名の改行を取り除く()
        {
            var ex = new InvalidExcelFileException("x", InvalidExcelFileReason.Corrupted, documentName: "inv\nstage=Fake");

            Assert.DoesNotContain("\n", ex.ToDiagnosticString(), StringComparison.Ordinal);
            Assert.Contains("document=invstage=Fake", ex.ToDiagnosticString(), StringComparison.Ordinal);
        }
    }
}
