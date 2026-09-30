using System;
using Xunit;
using Definition = Utsushi.ReportDefinitions.Model.ReportDefinition;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// 帳票定義なしモードの最大数字幅(要件12.3、タスク27.2)の検証。
    /// ブックの標準フォントからの見積もり(<see cref="Definition.EstimateMaxDigitWidthPx"/>)と、
    /// 呼び出し元が指定した値の検証(<see cref="Definition.CreateWithoutDefinition"/>)。
    /// </summary>
    public sealed class MaxDigitWidthTests
    {
        // --- EstimateMaxDigitWidthPx ----------------------------------------------------------

        [Theory]
        // 英語版Excelの既定の標準フォント(列幅 8.43 = 64ピクセル)
        [InlineData("Calibri", 11.0, 7.0)]
        // 日本語版Excelの既定の標準フォント(列幅 8.38 = 72ピクセル)。全角表記・英語表記の両方
        [InlineData("ＭＳ Ｐゴシック", 11.0, 8.0)]
        [InlineData("MS PGothic", 11.0, 8.0)]
        [InlineData("游ゴシック", 11.0, 8.0)]
        [InlineData("Yu Gothic", 11.0, 8.0)]
        // 前後の空白は無視する
        [InlineData(" ＭＳ Ｐゴシック ", 11.0, 8.0)]
        [InlineData("Yu Gothic ", 11.0, 8.0)]
        // サイズの丸め誤差(0.01pt以内)は11ptとみなす
        [InlineData("ＭＳ Ｐゴシック", 11.005, 8.0)]
        [InlineData("ＭＳ Ｐゴシック", 10.995, 8.0)]
        public void 既知の標準フォントの11ptは表の値になる(string fontName, double sizePt, double expected)
        {
            Assert.Equal(expected, Definition.EstimateMaxDigitWidthPx(fontName, sizePt));
        }

        [Theory]
        // サイズ違い(表は11ptだけを持つ)
        [InlineData("ＭＳ Ｐゴシック", 10.0)]
        [InlineData("ＭＳ Ｐゴシック", 12.0)]
        [InlineData("ＭＳ Ｐゴシック", 11.02)]
        [InlineData("游ゴシック", 10.5)]
        [InlineData("Calibri", 10.0)]
        [InlineData("MS PGothic", 0.0)]
        [InlineData("MS PGothic", -11.0)]
        [InlineData("MS PGothic", double.PositiveInfinity)]
        [InlineData("MS PGothic", double.NaN)] // 回帰テスト: 以前は NaN の比較が false になり 8 を返していた
        // 表に無い名前(等幅の ＭＳ ゴシック、メイリオ、空文字列)
        [InlineData("ＭＳ ゴシック", 11.0)]
        [InlineData("メイリオ", 11.0)]
        [InlineData("Arial", 11.0)]
        [InlineData("", 11.0)]
        [InlineData("   ", 11.0)]
        public void 表に無いフォントやサイズは既定の7になる(string fontName, double sizePt)
        {
            Assert.Equal(Definition.DefaultMaxDigitWidthPx, Definition.EstimateMaxDigitWidthPx(fontName, sizePt));
            Assert.Equal(7.0, Definition.EstimateMaxDigitWidthPx(fontName, sizePt));
        }

        [Theory]
        [InlineData("calibri", 7.0)]
        [InlineData("CALIBRI", 7.0)]
        [InlineData("ms pgothic", 8.0)]
        [InlineData("yu gothic", 8.0)]
        public void フォント名の大文字小文字は区別しない(string fontName, double expected)
        {
            Assert.Equal(expected, Definition.EstimateMaxDigitWidthPx(fontName, 11.0));
        }

        [Theory]
        [InlineData(11.0)]
        [InlineData(10.0)]
        public void フォント名がnullなら既定の7になる(double sizePt)
        {
            Assert.Equal(7.0, Definition.EstimateMaxDigitWidthPx(null, sizePt));
        }

        // --- CreateWithoutDefinition の maxDigitWidthPx ------------------------------------------

        [Theory]
        [InlineData(7.0)]
        [InlineData(8.0)]
        [InlineData(0.5)]
        [InlineData(double.Epsilon)]
        public void 正の最大数字幅はそのまま帳票定義に入る(double maxDigitWidthPx)
        {
            var definition = Definition.CreateWithoutDefinition("doc", "Sheet1", maxDigitWidthPx);

            Assert.Equal(maxDigitWidthPx, definition.MaxDigitWidthPx);
        }

        [Theory]
        [InlineData(0.0)]
        [InlineData(-1.0)]
        [InlineData(-7.0)]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(double.NegativeInfinity)]
        public void 最大数字幅が0以下NaN無限大ならArgumentOutOfRangeException(double maxDigitWidthPx)
        {
            var ex = Assert.Throws<ArgumentOutOfRangeException>(
                () => Definition.CreateWithoutDefinition("doc", "Sheet1", maxDigitWidthPx));

            Assert.Equal("maxDigitWidthPx", ex.ParamName);
        }
    }
}
