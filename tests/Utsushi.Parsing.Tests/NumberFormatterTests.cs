using System;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests;

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

    private static double ToSerial(DateTime value) => (value - new DateTime(1899, 12, 30)).TotalDays;
}
