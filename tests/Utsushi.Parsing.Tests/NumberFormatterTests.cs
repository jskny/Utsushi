using System;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 数値書式の適用(要件1.2 の「数値書式」)の検証。
    /// 自社帳票が使う範囲(金額・数量・日付・パーセント)に限定したサブセット実装の確認。
    /// </summary>
    public sealed class NumberFormatterTests
    {
        [Theory]
        [InlineData(1234.5, null, "1234.5")]
        [InlineData(1234.5, "General", "1234.5")]
        [InlineData(1234.0, "0", "1234")]
        [InlineData(1234.567, "0.00", "1234.57")]
        [InlineData(1234567.0, "#,##0", "1,234,567")]
        [InlineData(1234567.891, "#,##0.00", "1,234,567.89")]
        [InlineData(0.153, "0%", "15%")]
        [InlineData(0.1534, "0.00%", "15.34%")]
        public void 基本的な数値書式を適用する(double value, string? format, string expected)
        {
            Assert.Equal(expected, NumberFormatter.FormatNumber(value, format));
        }

        [Theory]
        [InlineData(320000.0, "\"¥\"#,##0", "¥320,000")]
        [InlineData(320000.0, "[$¥-411]#,##0", "¥320,000")]
        [InlineData(1500.0, "#,##0\"円\"", "1,500円")]
        public void 通貨記号つきの書式を適用する(double value, string format, string expected)
        {
            Assert.Equal(expected, NumberFormatter.FormatNumber(value, format));
        }

        [Fact]
        public void 負数セクションが指定されていれば使う()
        {
            Assert.Equal("1,000", NumberFormatter.FormatNumber(1000, "#,##0;(#,##0)"));
            Assert.Equal("(1,000)", NumberFormatter.FormatNumber(-1000, "#,##0;(#,##0)"));
        }

        [Fact]
        public void ゼロセクションが指定されていれば使う()
        {
            Assert.Equal("-", NumberFormatter.FormatNumber(0, "#,##0;-#,##0;\"-\""));
        }

        [Fact]
        public void 色指定は表示に影響しない()
        {
            Assert.Equal("1,000", NumberFormatter.FormatNumber(1000, "[Red]#,##0"));
        }

        [Theory]
        [InlineData(1234567.0, "#,##0,", "1,235")]
        [InlineData(1234567.0, "#,##0.0,", "1,234.6")]
        [InlineData(1234567890.0, "#,##0.00,,", "1,234.57")]
        public void 末尾のカンマは小数点の有無によらず桁区切りスケーリングになる(double value, string format, string expected)
        {
            Assert.Equal(expected, NumberFormatter.FormatNumber(value, format));
        }

        [Theory]
        [InlineData(2026, 4, 1, "yyyy/mm/dd", "2026/04/01")]
        [InlineData(2026, 4, 1, "yyyy\"年\"m\"月\"d\"日\"", "2026年4月1日")]
        [InlineData(2026, 12, 25, "yy/m/d", "26/12/25")]
        [InlineData(2026, 4, 30, "m/d", "4/30")]
        public void 日付書式を適用する(int year, int month, int day, string format, string expected)
        {
            var serial = ToSerial(new DateTime(year, month, day));

            Assert.Equal(expected, NumberFormatter.FormatNumber(serial, format));
        }

        [Fact]
        public void 時刻書式でmを分として解釈する()
        {
            var serial = ToSerial(new DateTime(2026, 4, 1, 13, 5, 9));

            // h の直後の m は「分」
            Assert.Equal("13:05", NumberFormatter.FormatNumber(serial, "h:mm"));
            Assert.Equal("13:05:09", NumberFormatter.FormatNumber(serial, "h:mm:ss"));

            // s の直前の m も「分」
            Assert.Equal("05:09", NumberFormatter.FormatNumber(serial, "mm:ss"));

            // 日付の文脈では「月」
            Assert.Equal("04/01", NumberFormatter.FormatNumber(serial, "mm/dd"));
        }

        [Fact]
        public void AMPM指定では12時間表記になる()
        {
            var afternoon = ToSerial(new DateTime(2026, 4, 1, 13, 5, 0));
            var morning = ToSerial(new DateTime(2026, 4, 1, 9, 5, 0));

            Assert.Equal("1:05 PM", NumberFormatter.FormatNumber(afternoon, "h:mm AM/PM"));
            Assert.Equal("9:05 AM", NumberFormatter.FormatNumber(morning, "h:mm AM/PM"));
        }

        [Fact]
        public void 日付とみなすかどうかを書式から判定する()
        {
            Assert.True(NumberFormatter.IsDateTimeFormat("yyyy/mm/dd"));
            Assert.True(NumberFormatter.IsDateTimeFormat("h:mm:ss"));
            Assert.False(NumberFormatter.IsDateTimeFormat("#,##0"));
            Assert.False(NumberFormatter.IsDateTimeFormat("\"¥\"#,##0"));

            // 引用符内の 'd' や 'm' は書式指定子ではない
            Assert.False(NumberFormatter.IsDateTimeFormat("#,##0\"md\""));

            // [Red] の 'd' も書式指定子ではない
            Assert.False(NumberFormatter.IsDateTimeFormat("[Red]#,##0"));
        }

        [Fact]
        public void シリアル値1は1900年1月1日になる()
        {
            // Excel は 1900年をうるう年とみなす既知の不具合があるため、60未満は補正が必要。
            Assert.Equal(new DateTime(1900, 1, 1), NumberFormatter.FromSerial(1));
            Assert.Equal(new DateTime(1900, 3, 1), NumberFormatter.FromSerial(61));
        }

        [Fact]
        public void 解釈できない書式はGeneral相当にフォールバックする()
        {
            // 指数表記・分数表記は対象外。例外にせず値を返す。
            var exponent = NumberFormatter.FormatNumber(1234.5, "0.00E+00");
            var fraction = NumberFormatter.FormatNumber(1.5, "# ?/?");

            Assert.Equal("1234.5", exponent);
            Assert.Equal("1.5", fraction);
        }

        [Theory]
        [InlineData(1.0625, "[h]:mm:ss", "25:30:00")]
        [InlineData(0.5, "[h]:mm", "12:00")]
        [InlineData(3.0, "[hh]:mm", "72:00")]
        [InlineData(1.0625, "[mm]:ss", "1530:00")]
        [InlineData(0.0625, "[s]", "5400")]
        public void 経過時間は24時間を超えて数える(double serial, string format, string expected)
        {
            Assert.Equal(expected, NumberFormatter.FormatNumber(serial, format));
        }

        [Fact]
        public void 組み込み書式46は経過時間になる()
        {
            Assert.Equal("25:30:00", NumberFormatter.FormatNumber(1.0625, NumberFormatter.GetBuiltInFormat(46)));
        }

        [Fact]
        public void 経過時間以外の角括弧は日付書式でも読み飛ばす()
        {
            var serial = ToSerial(new DateTime(2026, 4, 1));

            Assert.Equal("2026/4/1", NumberFormatter.FormatNumber(serial, "[Red]yyyy/m/d"));
            Assert.Equal("2026/4/1", NumberFormatter.FormatNumber(serial, "[$-411]yyyy/m/d"));
            Assert.True(NumberFormatter.IsDateTimeFormat("[h]"));
            Assert.False(NumberFormatter.IsDateTimeFormat("[$-411]#,##0"));
        }

        [Theory]
        [InlineData(2024, 4, 1, "[$-411]ggge\"年\"m\"月\"d\"日\"", "令和6年4月1日")]
        [InlineData(2024, 4, 1, "[$-411]gge\"年\"", "令6年")]
        [InlineData(2024, 4, 1, "[$-411]ge.m.d", "R6.4.1")]
        [InlineData(2024, 4, 1, "[$-411]gggee\"年\"", "令和06年")]
        [InlineData(2019, 5, 1, "ggge\"年\"", "令和1年")]
        [InlineData(2019, 4, 30, "ggge\"年\"", "平成31年")]
        [InlineData(1989, 1, 8, "ggge\"年\"", "平成1年")]
        [InlineData(1989, 1, 7, "ggge\"年\"", "昭和64年")]
        [InlineData(1926, 12, 25, "ge", "S1")]
        [InlineData(1926, 12, 24, "ge", "T15")]
        [InlineData(1912, 7, 30, "gge", "大1")]
        [InlineData(1912, 7, 29, "gge", "明45")]
        [InlineData(1900, 1, 1, "ggge", "明治33")]
        public void 和暦の元号と年を表示する(int year, int month, int day, string format, string expected)
        {
            var serial = ToSerial(new DateTime(year, month, day));

            Assert.Equal(expected, NumberFormatter.FormatNumber(serial, format));
        }

        [Fact]
        public void 和暦の組み込み書式を適用する()
        {
            var serial = 45383.0; // 2024/4/1

            Assert.Equal("令和6年4月1日", NumberFormatter.FormatNumber(serial, NumberFormatter.GetBuiltInFormat(28)));
            Assert.Equal("R6.4.1", NumberFormatter.FormatNumber(serial, NumberFormatter.GetBuiltInFormat(27)));
            Assert.Equal("2024年4月1日", NumberFormatter.FormatNumber(serial, NumberFormatter.GetBuiltInFormat(31)));
        }

        [Fact]
        public void 解釈できない日付書式指定子はGeneral相当にフォールバックする()
        {
            // b(仏暦)は未対応。書式文字をそのまま出さない。
            Assert.Equal("45383", NumberFormatter.FormatNumber(45383, "bbbb/m/d"));

            // 明治(1868/1/1)より前の日付は和暦で表せない。
            Assert.Equal("-20000", NumberFormatter.FormatNumber(-20000, "ggge"));

            // 負の経過時間も表示できない。
            Assert.Equal("-0.5", NumberFormatter.FormatNumber(-0.5, "[h]:mm"));
        }

        [Theory]
        [InlineData(123.0, "@", "123")]
        [InlineData(1234.5, "@", "1234.5")]
        [InlineData(1234.0, "#,##0;-#,##0;0;@", "1,234")]
        [InlineData(-5.0, "0;@", "-5")]
        public void 文字列書式の数値はGeneralで表示する(double value, string format, string expected)
        {
            Assert.Equal(expected, NumberFormatter.FormatNumber(value, format));
        }

        [Fact]
        public void 組み込み書式49の数値はGeneralで表示する()
        {
            Assert.Equal("123", NumberFormatter.FormatNumber(123, NumberFormatter.GetBuiltInFormat(49)));
        }

        [Theory]
        [InlineData(-1000.0, "\"¥\"#,##0", "-¥1,000")]
        [InlineData(-1000.0, "[$¥-411]#,##0", "-¥1,000")]
        [InlineData(-1500.0, "#,##0\"円\"", "-1,500円")]
        [InlineData(-0.5, ".00", "-.50")]
        [InlineData(-0.153, "0%", "-15%")]
        public void セクションが1つの書式では負号を出力全体の先頭に付ける(double value, string format, string expected)
        {
            Assert.Equal(expected, NumberFormatter.FormatNumber(value, format));
        }

        [Fact]
        public void セクションが複数の書式では負号を付けない()
        {
            Assert.Equal("▲1,000", NumberFormatter.FormatNumber(-1000, "#,##0;\"▲\"#,##0"));
            Assert.Equal("¥-1,000", NumberFormatter.FormatNumber(-1000, "\"¥\"#,##0;\"¥\"\\-#,##0"));
        }

        [Theory]
        [InlineData(1.5, "0.0#", "1.5")]
        [InlineData(1.25, "0.0#", "1.25")]
        [InlineData(1.0, "0.0#", "1.0")]
        [InlineData(1.0, "0.##", "1.")]
        [InlineData(1.5, "0.##", "1.5")]
        [InlineData(1234.5, "#,##0.##", "1,234.5")]
        [InlineData(1.5, "0.0?", "1.5 ")]
        [InlineData(1.25, "0.00", "1.25")]
        public void 小数部の井桁は不要な0を出さない(double value, string format, string expected)
        {
            Assert.Equal(expected, NumberFormatter.FormatNumber(value, format));
        }

        [Theory]
        [InlineData(14, "2026/4/1")]
        [InlineData(22, "2026/4/1 13:05")]
        [InlineData(31, "2026年4月1日")]
        [InlineData(32, "13時05分")]
        [InlineData(33, "13時05分09秒")]
        [InlineData(34, "2026年4月")]
        [InlineData(35, "4月1日")]
        public void 日付の組み込み書式は日本語版Excelの表示に合わせる(int numberFormatId, string expected)
        {
            var serial = ToSerial(new DateTime(2026, 4, 1, 13, 5, 9));

            Assert.Equal(expected, NumberFormatter.FormatNumber(serial, NumberFormatter.GetBuiltInFormat(numberFormatId)));
        }

        [Fact]
        public void 通貨の組み込み書式は円記号で表示する()
        {
            Assert.Equal("¥1,000", NumberFormatter.FormatNumber(1000, NumberFormatter.GetBuiltInFormat(5)));
            Assert.Equal("¥-1,000", NumberFormatter.FormatNumber(-1000, NumberFormatter.GetBuiltInFormat(5)));
        }

        [Fact]
        public void 秒の小数部を表示する()
        {
            var serial = ToSerial(new DateTime(2026, 4, 1, 13, 5, 9, 250));

            Assert.Equal("05:09.2", NumberFormatter.FormatNumber(serial, NumberFormatter.GetBuiltInFormat(47)));
            Assert.Equal("05:09.25", NumberFormatter.FormatNumber(serial, "mm:ss.00"));
        }

        [Fact]
        public void 引用符で区切った時と分でもmを分として解釈する()
        {
            var serial = ToSerial(new DateTime(2026, 4, 1, 13, 5, 9));

            Assert.Equal("13時05分", NumberFormatter.FormatNumber(serial, "h\"時\"mm\"分\""));
            Assert.Equal("4月1日", NumberFormatter.FormatNumber(serial, "m\"月\"d\"日\""));
        }

        [Fact]
        public void 同じ書式を並行して使っても同じ結果になる()
        {
            // 解析結果のキャッシュは複数スレッドから共有される。
            var results = new string[1000];
            System.Threading.Tasks.Parallel.For(0, results.Length, i =>
            {
                results[i] = NumberFormatter.FormatNumber(i * 1000.5, "\"¥\"#,##0.0#;\"▲\"#,##0");
            });

            for (var i = 0; i < results.Length; i++)
            {
                Assert.Equal(NumberFormatter.FormatNumber(i * 1000.5, "\"¥\"#,##0.0#;\"▲\"#,##0"), results[i]);
            }

            Assert.Equal("¥1,000.5", results[1]);
        }

        [Fact]
        public void 上限を超える数の書式コードでも正しく整形する()
        {
            // キャッシュの件数上限を超えても、都度解析して同じ結果を返す。
            for (var i = 0; i < 1100; i++)
            {
                var format = "#,##0\"" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\"";
                Assert.Equal("1,000" + i.ToString(System.Globalization.CultureInfo.InvariantCulture), NumberFormatter.FormatNumber(1000, format));
            }
        }

        private static double ToSerial(DateTime value) => (value - new DateTime(1899, 12, 30)).TotalDays;
    }
}
