using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout.Model
{
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

    /// <summary>
    /// 画像1枚の描画(要件9)。他のセル内容(背景・罫線・文字)より最前面に描画される。
    /// </summary>
    /// <param name="Rect">配置先の矩形(ページ左上原点、ポイント単位、余白・拡大縮小適用済み)。</param>
    /// <param name="Data">画像のバイナリ。</param>
    /// <param name="ContentType">MIMEタイプ(例: <c>"image/png"</c>)。</param>
    public sealed record ImageCommand(RectPt Rect, byte[] Data, string ContentType) : DrawCommand;

    /// <summary>
    /// 図形1つの描画(要件10)。他のセル内容(背景・罫線・文字)より最前面に描画される。
    /// </summary>
    /// <param name="Rect">配置先の矩形(ページ左上原点、ポイント単位、余白・拡大縮小適用済み)。</param>
    /// <param name="Preset">プリセットジオメトリの種別。</param>
    /// <param name="AdjustmentValues">
    /// プリセットごとに定めた順序の調整ガイド値。ファイルに指定が無い位置は<see cref="double.NaN"/>で、
    /// Renderingレイヤーがその位置のECMA-376既定値を補う。
    /// </param>
    /// <param name="RotationDegrees"><see cref="Rect"/>の中心を軸とした回転角(度、時計回り)。</param>
    /// <param name="Fill">塗りつぶし。<c>null</c>は塗りつぶし無し。</param>
    /// <param name="Outline">枠線。<c>null</c>は枠線無し。</param>
    /// <param name="TextLines">
    /// 図形内テキストの各行(折り返し・配置は確定済み)。座標は回転前のローカル座標であり、
    /// 回転の適用はRenderingレイヤーの責務。
    /// </param>
    public sealed record ShapeCommand(
        RectPt Rect,
        ShapePresetType Preset,
        IReadOnlyList<double> AdjustmentValues,
        double RotationDegrees,
        ShapeFill? Fill,
        ShapeOutline? Outline,
        IReadOnlyList<ShapeTextLine> TextLines) : DrawCommand;

    /// <summary>図形内テキストの1行(要件10.4)。座標は<see cref="ShapeCommand.Rect"/>を基準とした、回転前のローカル座標。</summary>
    public sealed record ShapeTextLine(PointPt Origin, string Text, FontStyle Font, TextAnchor Anchor);

    /// <summary>
    /// 接続線1本の描画(要件10.9)。塗りつぶし・テキストを持たない。
    /// </summary>
    /// <param name="Rect">配置先の矩形(ページ左上原点、ポイント単位、余白・拡大縮小適用済み)。</param>
    /// <param name="Preset">接続線のプリセット種別(直線・カギ線・曲線)。</param>
    /// <param name="RotationDegrees"><see cref="Rect"/>の中心を軸とした回転角(度、時計回り)。</param>
    /// <param name="FlipHorizontal">左右反転の有無(経路の折れ/曲がる向きを決める)。</param>
    /// <param name="FlipVertical">上下反転の有無(経路の折れ/曲がる向きを決める)。</param>
    /// <param name="Outline">枠線。<c>null</c>の場合、Renderingレイヤーが既定の黒い実線を補う。</param>
    /// <param name="ResolvedStart">
    /// 要件10.11の接続点解決に成功した場合の始点の絶対座標(ページ座標)。<c>null</c>の場合、
    /// Renderingレイヤーは<see cref="Rect"/>と<see cref="FlipHorizontal"/>/<see cref="FlipVertical"/>
    /// から始点を決める(要件10.9の既定動作)。
    /// </param>
    /// <param name="ResolvedEnd"><see cref="ResolvedStart"/>と同様の終点。</param>
    public sealed record ConnectorCommand(
        RectPt Rect,
        ConnectorPresetType Preset,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical,
        ShapeOutline? Outline,
        PointPt? ResolvedStart,
        PointPt? ResolvedEnd) : DrawCommand;

    /// <summary>
    /// グループ化された図形の展開結果(要件10.10)。<paramref name="Children"/>は
    /// グループの子座標空間からページ座標へ変換済みだが、グループ自身の回転は未適用であり、
    /// Renderingレイヤーが<paramref name="Center"/>を軸に<paramref name="RotationDegrees"/>だけ
    /// 回転させたうえで<paramref name="Children"/>を描画する(子要素個別の回転とは独立に合成する)。
    /// </summary>
    /// <param name="Center">グループ自身の配置矩形の中心(回転の軸)。</param>
    /// <param name="RotationDegrees">グループ全体の回転角(度、時計回り)。</param>
    /// <param name="Children">
    /// 子座標空間からページ座標へ変換済みの描画命令(出現順)。<see cref="ShapeCommand"/>/
    /// <see cref="ImageCommand"/>/<see cref="ConnectorCommand"/>/入れ子の<see cref="GroupCommand"/>のいずれか。
    /// </param>
    public sealed record GroupCommand(PointPt Center, double RotationDegrees, IReadOnlyList<DrawCommand> Children) : DrawCommand;

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
    /// <param name="Commands">描画命令(背景→罫線→テキスト→画像・図形(drawing.xmlの出現順)の順)。</param>
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
}
