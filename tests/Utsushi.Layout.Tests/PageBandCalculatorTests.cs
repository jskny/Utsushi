using System;
using System.Collections.Generic;
using System.Linq;
using Utsushi.Layout;
using Xunit;

namespace Utsushi.Layout.Tests;

/// <summary>
/// 改ページ計算の検証(要件3.2 手動改ページ / 要件3.3 自動改ページ、タスク5.3/5.4)。
/// </summary>
public sealed class PageBandCalculatorTests
{
    private static IReadOnlyList<IReadOnlyList<int>> Split(
        IEnumerable<int> indices, double sizeEach, double available, params int[] manualBreaks) =>
        PageBandCalculator.Split(
            indices.ToList(), _ => sizeEach, available, new HashSet<int>(manualBreaks));

    [Fact]
    public void 印字可能領域に収まるなら1ページになる()
    {
        var bands = Split(Enumerable.Range(1, 10), sizeEach: 10.0, available: 100.0);

        var band = Assert.Single(bands);
        Assert.Equal(Enumerable.Range(1, 10), band);
    }

    [Fact]
    public void 印字可能領域を超えたら自動的に分割する()
    {
        // 1要素10pt、領域35pt → 3要素ずつ
        var bands = Split(Enumerable.Range(1, 10), sizeEach: 10.0, available: 35.0);

        Assert.Equal(4, bands.Count);
        Assert.Equal(new[] { 1, 2, 3 }, bands[0]);
        Assert.Equal(new[] { 4, 5, 6 }, bands[1]);
        Assert.Equal(new[] { 7, 8, 9 }, bands[2]);
        Assert.Equal(new[] { 10 }, bands[3]);
    }

    [Fact]
    public void 境界ちょうどに収まる場合は次ページへ送らない()
    {
        // 1要素10pt、領域30pt → ちょうど3要素
        var bands = Split(Enumerable.Range(1, 6), sizeEach: 10.0, available: 30.0);

        Assert.Equal(2, bands.Count);
        Assert.Equal(new[] { 1, 2, 3 }, bands[0]);
        Assert.Equal(new[] { 4, 5, 6 }, bands[1]);
    }

    [Fact]
    public void 手動改ページは指定位置の手前で分割する()
    {
        var bands = Split(Enumerable.Range(1, 10), sizeEach: 1.0, available: 1000.0, manualBreaks: 4);

        Assert.Equal(2, bands.Count);
        Assert.Equal(new[] { 1, 2, 3 }, bands[0]);
        Assert.Equal(new[] { 4, 5, 6, 7, 8, 9, 10 }, bands[1]);
    }

    [Fact]
    public void 手動改ページが複数あればそれぞれで分割する()
    {
        var bands = Split(Enumerable.Range(1, 9), sizeEach: 1.0, available: 1000.0, manualBreaks: new[] { 4, 7 });

        Assert.Equal(3, bands.Count);
        Assert.Equal(new[] { 1, 2, 3 }, bands[0]);
        Assert.Equal(new[] { 4, 5, 6 }, bands[1]);
        Assert.Equal(new[] { 7, 8, 9 }, bands[2]);
    }

    [Fact]
    public void 先頭位置の手動改ページは空ページを作らない()
    {
        var bands = Split(Enumerable.Range(1, 3), sizeEach: 1.0, available: 1000.0, manualBreaks: 1);

        var band = Assert.Single(bands);
        Assert.Equal(new[] { 1, 2, 3 }, band);
    }

    [Fact]
    public void 手動改ページと自動改ページは併用できる()
    {
        // 1要素10pt、領域25pt(=2要素)、さらに5の手前で手動改ページ
        var bands = Split(Enumerable.Range(1, 8), sizeEach: 10.0, available: 25.0, manualBreaks: 5);

        Assert.Equal(new[] { 1, 2 }, bands[0]);
        Assert.Equal(new[] { 3, 4 }, bands[1]);
        Assert.Equal(new[] { 5, 6 }, bands[2]);
        Assert.Equal(new[] { 7, 8 }, bands[3]);
    }

    [Fact]
    public void 単独で印字可能領域を超える行はそれだけで1ページになる()
    {
        // 50pt の要素が領域30ptに入らないが、空ページは作らずその要素だけのページにする
        var sizes = new Dictionary<int, double> { [1] = 10.0, [2] = 50.0, [3] = 10.0 };
        var bands = PageBandCalculator.Split(
            new[] { 1, 2, 3 }, i => sizes[i], 30.0, Array.Empty<int>());

        Assert.Equal(3, bands.Count);
        Assert.Equal(new[] { 1 }, bands[0]);
        Assert.Equal(new[] { 2 }, bands[1]);
        Assert.Equal(new[] { 3 }, bands[2]);
    }

    [Fact]
    public void 対象が空でも空の帯を1つ返す()
    {
        var bands = PageBandCalculator.Split(Array.Empty<int>(), _ => 1.0, 100.0, Array.Empty<int>());

        Assert.Single(bands);
        Assert.Empty(bands[0]);
    }
}
