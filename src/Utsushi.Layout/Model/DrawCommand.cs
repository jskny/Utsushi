using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout.Model;

/// <summary>
/// 1ページに対する描画命令。座標・寸法はすべてページ左上原点のポイント単位で確定済みであり、
/// 拡大縮小率も適用済みである(Rendering レイヤーは追加の座標変換を行わない)。
/// </summary>
public abstract record DrawCommand;

/// <summary>矩形の塗りつぶし(セル背景)。</summary>
public sealed record FillRectCommand(RectPt Rect, ArgbColor Color) : DrawCommand;

/// <summary>線分(罫線)。</summary>
/// <param name="From">始点。</param>
/// <param name="To">終点。</param>
/// <param name="Color">線色。</param>
/// <param name="WidthPt">線幅(ポイント)。</param>
/// <param name="Dash">破線パターン。</param>
public sealed record LineCommand(PointPt From, PointPt To, ArgbColor Color, double WidthPt, LineDashStyle Dash)
    : DrawCommand;

/// <summary>
/// テキスト1行の描画。
/// </summary>
/// <param name="Origin">
/// 描画開始位置。X は <paramref name="Anchor"/> の基準点、Y はベースライン。
/// </param>
/// <param name="Text">描画する文字列。</param>
/// <param name="Font">フォント(サイズは縮小表示・拡大縮小率の適用後)。</param>
/// <param name="Anchor">X座標をどう解釈するか。</param>
/// <param name="ClipRect">描画をこの矩形で切り取る。切り取り不要の場合は null。</param>
public sealed record TextCommand(
    PointPt Origin,
    string Text,
    FontStyle Font,
    TextAnchor Anchor,
    RectPt? ClipRect) : DrawCommand;

/// <summary>テキストのX座標の解釈。</summary>
public enum TextAnchor
{
    /// <summary>X は文字列の左端。</summary>
    Left = 0,

    /// <summary>X は文字列の中心。</summary>
    Center,

    /// <summary>X は文字列の右端。</summary>
    Right,
}

/// <summary>罫線の破線パターン。</summary>
public enum LineDashStyle
{
    Solid = 0,
    Dot,
    Dash,
    DashDot,
    DashDotDot,
}

/// <summary>
/// 1ページ分のレイアウト結果。
/// </summary>
/// <param name="Paper">用紙。</param>
/// <param name="Orientation">印刷の向き。</param>
/// <param name="WidthPt">ページ幅(向き適用後、ポイント)。</param>
/// <param name="HeightPt">ページ高さ(向き適用後、ポイント)。</param>
/// <param name="Commands">描画命令(背景→罫線→テキストの順)。</param>
/// <param name="PageNumber">1始まりのページ番号。</param>
/// <param name="RowRange">このページが表示する本文行の範囲(診断・テスト用)。</param>
/// <param name="ColumnRange">このページが表示する本文列の範囲(診断・テスト用)。</param>
/// <param name="ScaleFactor">適用された拡大縮小率(1.0 = 等倍。座標には適用済み)。</param>
public sealed record PageLayout(
    PaperSize Paper,
    PageOrientation Orientation,
    double WidthPt,
    double HeightPt,
    IReadOnlyList<DrawCommand> Commands,
    int PageNumber,
    (int First, int Last) RowRange,
    (int First, int Last) ColumnRange,
    double ScaleFactor);

/// <summary>ページ分割・座標計算済みのレイアウト結果。</summary>
/// <param name="Pages">印刷順に並んだページ。</param>
/// <param name="ReportCode">帳票コード(診断用)。</param>
/// <param name="SheetName">シート名(診断用)。</param>
public sealed record PagedLayout(IReadOnlyList<PageLayout> Pages, string ReportCode, string SheetName)
{
    public int PageCount => Pages.Count;
}
