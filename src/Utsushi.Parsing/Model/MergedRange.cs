using Utsushi.Core;

namespace Utsushi.Parsing.Model;

/// <summary>結合セルの範囲。</summary>
public sealed record MergedRange(CellRange Range)
{
    /// <summary>結合範囲の左上セル。値・書式はこのセルのものを使う。</summary>
    public CellAddress Anchor => Range.TopLeft;

    public bool Contains(CellAddress address) => Range.Contains(address);
}
