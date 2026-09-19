using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Parsing.Model;

/// <summary>印刷の向き。</summary>
public enum PageOrientation
{
    /// <summary>Excel上「既定」。用紙の既定の向き(縦)として扱う。</summary>
    Default = 0,
    Portrait,
    Landscape,
}

/// <summary>複数ページの印刷順序(要件3.5)。</summary>
public enum PageOrder
{
    /// <summary>上から下、そのあと右へ(Excel既定)。</summary>
    DownThenOver = 0,

    /// <summary>左から右、そのあと下へ。</summary>
    OverThenDown,
}

/// <summary>
/// 用紙サイズ。OOXML の <c>pageSetup/@paperSize</c> コードと、ポイント単位の実寸(縦向き時)を保持する。
/// </summary>
public sealed record PaperSize(int Code, string Name, double WidthPt, double HeightPt)
{
    public static PaperSize A4 { get; } = FromMillimeters(9, "A4", 210, 297);

    public static PaperSize FromMillimeters(int code, string name, double widthMm, double heightMm) =>
        new(code, name, Units.MillimetersToPoints(widthMm), Units.MillimetersToPoints(heightMm));

    /// <summary>向きを適用した実寸(幅, 高さ)をポイントで返す。</summary>
    public (double WidthPt, double HeightPt) GetSize(PageOrientation orientation) =>
        orientation == PageOrientation.Landscape ? (HeightPt, WidthPt) : (WidthPt, HeightPt);
}

/// <summary>ページ余白(ポイント単位)。</summary>
public sealed record PageMargins(
    double LeftPt,
    double RightPt,
    double TopPt,
    double BottomPt,
    double HeaderPt,
    double FooterPt)
{
    /// <summary>Excel の既定余白(左右 0.7in、上下 0.75in、ヘッダー/フッター 0.3in)。</summary>
    public static PageMargins Default { get; } = new(
        Units.InchesToPoints(0.7),
        Units.InchesToPoints(0.7),
        Units.InchesToPoints(0.75),
        Units.InchesToPoints(0.75),
        Units.InchesToPoints(0.3),
        Units.InchesToPoints(0.3));
}

/// <summary>
/// 印刷タイトル(2ページ目以降に繰り返す行/列)。要件3.4。
/// </summary>
/// <param name="FirstRow">繰り返す先頭行(1始まり)。行の繰り返しが無い場合は null。</param>
/// <param name="LastRow">繰り返す末尾行(1始まり)。</param>
/// <param name="FirstColumn">繰り返す先頭列(1始まり)。列の繰り返しが無い場合は null。</param>
/// <param name="LastColumn">繰り返す末尾列(1始まり)。</param>
public sealed record PrintTitles(int? FirstRow, int? LastRow, int? FirstColumn, int? LastColumn)
{
    public static PrintTitles None { get; } = new(null, null, null, null);

    public bool HasRows => FirstRow.HasValue && LastRow.HasValue;

    public bool HasColumns => FirstColumn.HasValue && LastColumn.HasValue;
}

/// <summary>
/// 拡大縮小設定。<see cref="FitToWidth"/>/<see cref="FitToHeight"/> が指定されている場合は
/// 固定倍率(<see cref="ScalePercent"/>)ではなくページ数に合わせる設定が優先される。
/// </summary>
public sealed record PageScaling(int ScalePercent, int? FitToWidth, int? FitToHeight)
{
    /// <summary>等倍(100%)。</summary>
    public static PageScaling Normal { get; } = new(100, null, null);

    /// <summary>「次のページ数に合わせて印刷」が有効かどうか。</summary>
    public bool IsFitToPage => FitToWidth.HasValue || FitToHeight.HasValue;

    /// <summary>固定倍率を係数(1.0 = 等倍)で返す。</summary>
    public double ScaleFactor => ScalePercent / 100.0;
}

/// <summary>
/// シートのページ設定。要件1.3 が求める項目を保持する。
/// </summary>
public sealed record PageSetupModel(
    PaperSize Paper,
    PageOrientation Orientation,
    PageMargins Margins,
    PageScaling Scaling,
    IReadOnlyList<CellRange> PrintAreas,
    IReadOnlyList<int> ManualRowBreaks,
    IReadOnlyList<int> ManualColumnBreaks,
    PrintTitles PrintTitles,
    PageOrder PageOrder)
{
    /// <summary>ページ設定が未指定のシート向けの既定値。</summary>
    public static PageSetupModel Default { get; } = new(
        PaperSize.A4,
        PageOrientation.Portrait,
        PageMargins.Default,
        PageScaling.Normal,
        new List<CellRange>(),
        new List<int>(),
        new List<int>(),
        PrintTitles.None,
        PageOrder.DownThenOver);

    /// <summary>向きを適用した用紙の実寸(ポイント)。</summary>
    public (double WidthPt, double HeightPt) PaperSizePt => Paper.GetSize(Orientation);
}
