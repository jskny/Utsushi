using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// シートに配置された図形(要件10)。対応済みプリセット一覧に含まれない
    /// プリセットジオメトリは読み取り対象にせず、サポート外要素として扱う。
    /// </summary>
    /// <param name="Preset">プリセットジオメトリの種別。</param>
    /// <param name="AdjustmentValues">
    /// <c>a:avLst</c> のガイド値。<see cref="Preset"/> ごとに定めた順序で並ぶ(0〜1の比率)。
    /// 対応するガイドがファイルに無い場合は、そのプリセットのECMA-376既定値を用いるため空でよい。
    /// </param>
    /// <param name="RotationDegrees"><c>a:xfrm/@rot</c> から変換した回転角(度、時計回り)。</param>
    /// <param name="Fill">塗りつぶし。<c>null</c> は塗りつぶし無し(<c>a:noFill</c>)。</param>
    /// <param name="Outline">枠線。<c>null</c> は枠線無し。</param>
    /// <param name="Text">図形内テキスト。<c>null</c> は <c>xdr:txBody</c> 無し。</param>
    /// <param name="AnchorCell">アンカー左上セル。</param>
    /// <param name="AnchorOffset">アンカーセル左上からのオフセット(ポイント)。</param>
    /// <param name="Extent">図形の終端(サイズ)の決め方。</param>
    public sealed record ShapeModel(
        ShapePresetType Preset,
        IReadOnlyList<double> AdjustmentValues,
        double RotationDegrees,
        ShapeFill? Fill,
        ShapeOutline? Outline,
        ShapeTextBody? Text,
        CellAddress AnchorCell,
        PointPt AnchorOffset,
        AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

    /// <summary>
    /// 対応済みのプリセットジオメトリ(要件10.1補足)。自社帳票での実用上の必要性を
    /// 踏まえたキュレーション方式であり、ECMA-376の <c>ST_ShapeType</c> 全体には対応しない。
    /// </summary>
    public enum ShapePresetType
    {
        Rect,
        RoundRect,
        Ellipse,
        Triangle,
        RightArrow,
        LeftArrow,
        UpArrow,
        DownArrow,
        LeftRightArrow,
        UpDownArrow,
        WedgeRectCallout,
        WedgeRoundRectCallout,
        WedgeEllipseCallout,
    }

    /// <summary>図形の塗りつぶし。</summary>
    public abstract record ShapeFill;

    /// <summary><c>a:solidFill</c> 相当の単色塗りつぶし。</summary>
    public sealed record SolidShapeFill(ArgbColor Color) : ShapeFill;

    /// <summary>
    /// <c>a:gradFill</c> 相当の線形グラデーション。先頭・末尾の <c>a:gs</c> の色のみを
    /// 開始色・終了色として採用する近似実装(要件10.6)。
    /// </summary>
    /// <param name="AngleDegrees"><c>a:lin/@ang</c> から変換した角度(度)。指定が無ければ0(左から右)。</param>
    public sealed record LinearGradientShapeFill(ArgbColor StartColor, ArgbColor EndColor, double AngleDegrees) : ShapeFill;

    /// <summary>図形の枠線(<c>a:ln</c>)。</summary>
    public sealed record ShapeOutline(ArgbColor Color, double WidthPt);

    /// <summary>
    /// 図形内テキスト(<c>xdr:txBody</c>)。折り返しは行わず、段落・ランをそのまま保持する
    /// (折り返しはLayoutレイヤーの責務。design.md「Layout レイヤー」参照)。
    /// </summary>
    /// <param name="VAlign"><c>a:bodyPr/@anchor</c> から変換した垂直配置。</param>
    public sealed record ShapeTextBody(IReadOnlyList<ShapeTextParagraph> Paragraphs, VerticalAlignment VAlign);

    /// <summary>図形内テキストの1段落。</summary>
    /// <param name="HAlign"><c>a:pPr/@algn</c> から変換した水平配置。</param>
    public sealed record ShapeTextParagraph(IReadOnlyList<ShapeTextRun> Runs, HorizontalAlignment HAlign);

    /// <summary>図形内テキストの1ラン(<c>a:r</c>)。</summary>
    public sealed record ShapeTextRun(string Text, FontStyle Font);
}
