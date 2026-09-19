using Utsushi.Core;
using Utsushi.Layout;
using Xunit;

namespace Utsushi.Layout.Tests;

/// <summary>
/// 列幅(文字数単位)→ピクセル/ポイント換算の検証(要件4.5、タスク5.1)。
/// </summary>
public sealed class ExcelUnitConverterTests
{
    /// <summary>
    /// ECMA-376 の換算式は Truncate を2回含むため、素朴な width×MDW とは一致しない。
    /// 期待値は式 <c>Truncate(((256*w + Truncate(128/MDW))/256)*MDW)</c> から求めた値。
    /// </summary>
    /// <remarks>
    /// OOXML の <c>col/@width</c> は「表示できる文字数」そのものではなく、
    /// 左右のパディング(5px)を含めて文字数へ換算し直した値であることに注意。
    /// 例えば「1文字ぶん表示できる列」の保存値は 1.0 ではなく約 1.7109 になる。
    /// </remarks>
    [Theory]
    [InlineData(8.43, 7.0, 59)]        // Excel 標準の列幅
    [InlineData(1.0, 7.0, 7)]
    [InlineData(1.7109375, 7.0, 12)]   // 1文字を表示できる列(=7px + 5pxパディング)
    [InlineData(10.0, 7.0, 70)]
    [InlineData(12.0, 7.0, 84)]
    [InlineData(4.0, 7.0, 28)]
    public void 列幅を96dpiのピクセルへ換算する(double width, double maxDigitWidthPx, double expectedPixels)
    {
        var actual = ExcelUnitConverter.ColumnWidthToPixels(width, maxDigitWidthPx);

        Assert.Equal(expectedPixels, actual, precision: 6);
    }

    [Fact]
    public void 列幅0は幅0として扱う()
    {
        Assert.Equal(0.0, ExcelUnitConverter.ColumnWidthToPixels(0.0, 7.0));
        Assert.Equal(0.0, ExcelUnitConverter.ColumnWidthToPoints(0.0, 7.0));
    }

    [Fact]
    public void ピクセルからポイントへの換算は96dpi基準になる()
    {
        // 96px = 1インチ = 72pt
        Assert.Equal(72.0, Units.PixelsToPoints(96.0), precision: 9);

        // 標準列幅 8.43 文字 = 59px = 44.25pt
        Assert.Equal(44.25, ExcelUnitConverter.ColumnWidthToPoints(8.43, 7.0), precision: 9);
    }

    [Fact]
    public void 最大数字幅が変わると列幅の換算結果も変わる()
    {
        var narrow = ExcelUnitConverter.ColumnWidthToPixels(10.0, 6.0);
        var wide = ExcelUnitConverter.ColumnWidthToPixels(10.0, 8.0);

        Assert.True(narrow < wide, $"MDWが大きいほど広くなるはず (narrow={narrow}, wide={wide})");
    }

    [Fact]
    public void 最大数字幅が0以下なら例外になる()
    {
        Assert.Throws<System.ArgumentOutOfRangeException>(
            () => ExcelUnitConverter.ColumnWidthToPixels(10.0, 0.0));
    }

    [Theory]
    [InlineData(0, 0.0)]
    [InlineData(1, 15.75)]  // 3文字 × 7px = 21px = 15.75pt
    [InlineData(2, 31.5)]
    public void インデントは1段あたり標準フォント3文字分になる(int level, double expectedPt)
    {
        Assert.Equal(expectedPt, ExcelUnitConverter.IndentWidthToPoints(level, 7.0), precision: 9);
    }
}
