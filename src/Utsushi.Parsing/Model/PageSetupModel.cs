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
/// ページヘッダー/フッターの設定(要件3.7〜3.9)。
/// </summary>
/// <remarks>
/// 各文字列は Excel の書式コード(<c>&amp;L</c>/<c>&amp;C</c>/<c>&amp;R</c> によるセクション指定、
/// <c>&amp;P</c> ページ番号など)を含んだ生の値。解釈は Layout レイヤーが行う。
/// </remarks>
/// <param name="OddHeader">奇数ページ(既定)のヘッダー。</param>
/// <param name="OddFooter">奇数ページ(既定)のフッター。</param>
/// <param name="EvenHeader">偶数ページのヘッダー。<paramref name="DifferentOddEven"/> が true のときに使う。</param>
/// <param name="EvenFooter">偶数ページのフッター。</param>
/// <param name="FirstHeader">先頭ページのヘッダー。<paramref name="DifferentFirst"/> が true のときに使う。</param>
/// <param name="FirstFooter">先頭ページのフッター。</param>
/// <param name="DifferentOddEven">奇数/偶数ページで別の指定を使うかどうか。</param>
/// <param name="DifferentFirst">先頭ページだけ別の指定を使うかどうか。</param>
/// <param name="ScaleWithDocument">拡大縮小率をヘッダー/フッターにも適用するかどうか。</param>
public sealed record HeaderFooterModel(
    string? OddHeader,
    string? OddFooter,
    string? EvenHeader,
    string? EvenFooter,
    string? FirstHeader,
    string? FirstFooter,
    bool DifferentOddEven,
    bool DifferentFirst,
    bool ScaleWithDocument)
{
    /// <summary>ヘッダー/フッターなし。</summary>
    public static HeaderFooterModel None { get; } =
        new(null, null, null, null, null, null, false, false, true);

    /// <summary>ヘッダーもフッターも設定されていないかどうか。</summary>
    public bool IsEmpty =>
        string.IsNullOrEmpty(OddHeader) && string.IsNullOrEmpty(OddFooter)
        && string.IsNullOrEmpty(EvenHeader) && string.IsNullOrEmpty(EvenFooter)
        && string.IsNullOrEmpty(FirstHeader) && string.IsNullOrEmpty(FirstFooter);

    /// <summary>指定ページ(1始まり)に適用するヘッダーを返す。</summary>
    public string? GetHeader(int pageNumber) =>
        Select(pageNumber, FirstHeader, EvenHeader, OddHeader);

    /// <summary>指定ページ(1始まり)に適用するフッターを返す。</summary>
    public string? GetFooter(int pageNumber) =>
        Select(pageNumber, FirstFooter, EvenFooter, OddFooter);

    private string? Select(int pageNumber, string? first, string? even, string? odd)
    {
        // 「先頭ページのみ別指定」が優先。次に奇数/偶数の別指定。
        // 該当する指定が空の場合、Excel はそのページのヘッダー/フッターを表示しない。
        if (DifferentFirst && pageNumber == 1)
        {
            return first;
        }

        if (DifferentOddEven && pageNumber % 2 == 0)
        {
            return even;
        }

        return odd;
    }
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
    PageOrder PageOrder,
    HeaderFooterModel HeaderFooter)
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
        PageOrder.DownThenOver,
        HeaderFooterModel.None);

    /// <summary>向きを適用した用紙の実寸(ポイント)。</summary>
    public (double WidthPt, double HeightPt) PaperSizePt => Paper.GetSize(Orientation);
}
