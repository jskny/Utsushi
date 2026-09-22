using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// シートに配置された図形(要件10)。対応済みプリセット一覧に含まれない
    /// プリセットジオメトリは読み取り対象にせず、サポート外要素として扱う。
    /// </summary>
    /// <param name="Id">
    /// <c>NonVisualDrawingProperties/@id</c>。接続線の接続先解決(要件10.11)のために保持する。
    /// </param>
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
        uint Id,
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
        CloudCallout,
        Callout1,
        Callout2,
        Callout3,
        Star4,
        Star5,
        Star6,
        Star8,
        FlowChartProcess,
        FlowChartDecision,
        FlowChartTerminator,
        FlowChartInputOutput,
        FlowChartDocument,
        FlowChartPredefinedProcess,
        FlowChartConnector,
    }

    /// <summary>
    /// Layout(接続点解決、要件10.11の左右の接続点の補正)・Rendering(実際の描画)の両方が使う
    /// 図形ジオメトリの比率定数。<c>Utsushi.Rendering</c>は<c>Utsushi.Layout</c>に依存する
    /// 向きであり、Layoutから見てRendering内部の定数を直接参照することはできない
    /// (逆方向の参照は循環参照になりビルドできない。<c>.kiro/steering/structure.md</c>の
    /// レイヤー依存の一方向ルール)。そのため、双方が既に依存している
    /// <c>Utsushi.Parsing.Model</c>にこの定数を置く。<c>flowChartDocument</c>の波形の深さ比率
    /// (<c>DocumentWaveDepthRatio</c>)はLayoutから参照する必要が無いため、
    /// <c>Utsushi.Rendering.ShapeGeometryBuilder</c>内部の<c>private</c>定数のまま残す。
    /// </summary>
    public static class ShapeGeometryConstants
    {
        /// <summary>flowChartInputOutput(平行四辺形)の上下辺のずらし幅(矩形の幅に対する比率)。</summary>
        public const double InputOutputSkewRatio = 0.2;
    }

    /// <summary>図形の塗りつぶし。</summary>
    public abstract record ShapeFill;

    /// <summary><c>a:solidFill</c> 相当の単色塗りつぶし。</summary>
    public sealed record SolidShapeFill(ArgbColor Color) : ShapeFill;

    /// <summary><c>a:gsLst/a:gs</c> の1つ(位置・色)。<paramref name="Position"/> は0.0〜1.0。</summary>
    public sealed record GradientStop(double Position, ArgbColor Color);

    /// <summary>
    /// <c>a:gradFill</c>/<c>a:lin</c> 相当の線形グラデーション。3点以上のグラデーション
    /// ストップに対応する(要件10.6)。
    /// </summary>
    /// <param name="Stops"><c>a:gsLst</c> の全ストップ(2点以上)。位置の昇順とは限らない。</param>
    /// <param name="AngleDegrees"><c>a:lin/@ang</c> から変換した角度(度)。指定が無ければ0(左から右)。</param>
    public sealed record LinearGradientShapeFill(IReadOnlyList<GradientStop> Stops, double AngleDegrees) : ShapeFill;

    /// <summary>
    /// <c>a:gradFill</c>/<c>a:path[@path='circle']</c> 相当の放射状グラデーション(要件10.6)。
    /// <c>a:path[@path='rect'/'shape']</c> も近似としてこの型で表す。
    /// </summary>
    /// <param name="Stops"><c>a:gsLst</c> の全ストップ(2点以上)。</param>
    /// <param name="CenterFraction">
    /// 図形の矩形に対する放射の中心位置の割合(<c>a:fillToRect</c> 由来)。既定は中心(0.5, 0.5)。
    /// </param>
    public sealed record RadialGradientShapeFill(IReadOnlyList<GradientStop> Stops, PointPt CenterFraction) : ShapeFill;

    /// <summary>図形の枠線(<c>a:ln</c>)。接続線の枠線としても使う。</summary>
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

    /// <summary>
    /// 接続線(<c>xdr:cxnSp</c>。要件10.9)。図形(<c>xdr:sp</c>)とは別のOOXML要素であり、
    /// 塗りつぶし・テキストを持たない。
    /// </summary>
    /// <param name="Preset">対応済みの接続線プリセット。</param>
    /// <param name="RotationDegrees"><c>a:xfrm/@rot</c> から変換した回転角(度、時計回り)。</param>
    /// <param name="FlipHorizontal"><c>a:xfrm/@flipH</c>。経路の左右の向きを決める。</param>
    /// <param name="FlipVertical"><c>a:xfrm/@flipV</c>。経路の上下の向きを決める。</param>
    /// <param name="Outline">枠線。<c>null</c> の場合、描画時にExcelの既定(黒の実線1pt)を補う。</param>
    /// <param name="StartConnection">
    /// <c>a:stCxn</c>(始点の接続先)。要素が無ければ<c>null</c>。解決(参照先の矩形取得・
    /// 接続点座標の計算)はLayoutレイヤーの責務(要件10.11)。
    /// </param>
    /// <param name="EndConnection"><c>a:endCxn</c>(終点の接続先)。<see cref="StartConnection"/>と同様。</param>
    public sealed record ConnectorModel(
        ConnectorPresetType Preset,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical,
        ShapeOutline? Outline,
        ConnectionRef? StartConnection,
        ConnectionRef? EndConnection,
        CellAddress AnchorCell,
        PointPt AnchorOffset,
        AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

    /// <summary>
    /// 対応済みの接続線プリセット(要件10.9補足)。OOXMLの<c>bentConnectorN</c>等の
    /// 「N」は接続点の数ではなく折れ/曲がりの区分を表すため、本プロダクトでは
    /// 分かりやすさを優先し線分数(セグメント数)で名付ける。
    /// </summary>
    public enum ConnectorPresetType
    {
        /// <summary><c>straightConnector1</c>: 直線。</summary>
        Straight,

        /// <summary><c>bentConnector2</c>: 1回折れる(2辺)カギ線。</summary>
        Bent2Segment,

        /// <summary><c>bentConnector3</c>: 2回折れる(3辺)カギ線。</summary>
        Bent3Segment,

        /// <summary><c>curvedConnector2</c>: 1本の曲線。</summary>
        Curved2Segment,

        /// <summary><c>curvedConnector3</c>: 2本の曲線によるS字カーブ。</summary>
        Curved3Segment,
    }

    /// <summary>
    /// <c>a:stCxn</c>/<c>a:endCxn</c>(要件10.11)。<paramref name="ShapeId"/>は参照先の
    /// <c>NonVisualDrawingProperties/@id</c>と同じ値、<paramref name="SiteIndex"/>は
    /// 参照先の矩形上のどの接続点かを表す番号(<c>@idx</c>)。
    /// </summary>
    public sealed record ConnectionRef(uint ShapeId, uint SiteIndex);

    /// <summary>
    /// グループ化された図形(<c>xdr:grpSp</c>。要件10.10)。トップレベルの描画オブジェクトとして
    /// セルアンカーを持つが、内部の<see cref="Children"/>は独自の子座標空間
    /// (<see cref="ChildOffset"/>/<see cref="ChildExtent"/>)上の位置で決まる。
    /// </summary>
    /// <param name="Id">
    /// <c>NonVisualDrawingProperties/@id</c>。接続線の接続先解決(要件10.11)のために保持する。
    /// </param>
    /// <param name="ChildOffset"><c>a:chOff</c>(ポイント換算)。子要素の座標系の原点。</param>
    /// <param name="ChildExtent"><c>a:chExt</c>(ポイント換算)。X=幅、Y=高さ。</param>
    /// <param name="Children">直接の子要素(<c>drawing.xml</c>上の出現順)。</param>
    /// <param name="RotationDegrees"><c>a:xfrm/@rot</c> から変換したグループ自身の回転角(度)。</param>
    public sealed record GroupShapeModel(
        uint Id,
        PointPt ChildOffset,
        PointPt ChildExtent,
        IReadOnlyList<GroupChildModel> Children,
        double RotationDegrees,
        CellAddress AnchorCell,
        PointPt AnchorOffset,
        AnchorExtent Extent) : DrawingObjectModel(AnchorCell, AnchorOffset, Extent);

    /// <summary>
    /// グループ内の子要素の共通の位置決め情報。<see cref="LocalRect"/>は、それを含む
    /// グループの子座標空間(<c>a:chOff</c>/<c>a:chExt</c>)上の位置・サイズ(ポイント)であり、
    /// ページ座標への変換(平行移動+拡大縮小)はLayoutレイヤーが行う。
    /// </summary>
    public abstract record GroupChildModel(RectPt LocalRect);

    /// <summary>グループ内の図形。<see cref="Id"/>は接続線の接続先解決(要件10.11)のために保持する。</summary>
    public sealed record GroupChildShape(
        uint Id,
        RectPt LocalRect,
        ShapePresetType Preset,
        IReadOnlyList<double> AdjustmentValues,
        double RotationDegrees,
        ShapeFill? Fill,
        ShapeOutline? Outline,
        ShapeTextBody? Text) : GroupChildModel(LocalRect);

    /// <summary>グループ内の画像。<see cref="Id"/>は接続線の接続先解決(要件10.11)のために保持する。</summary>
    public sealed record GroupChildImage(
        uint Id, RectPt LocalRect, byte[] Data, string ContentType) : GroupChildModel(LocalRect);

    /// <summary>
    /// グループ内の接続線。<see cref="Id"/>は持たない(接続先として参照される対象ではないため。
    /// 要件10.11補足)。
    /// </summary>
    public sealed record GroupChildConnector(
        RectPt LocalRect,
        ConnectorPresetType Preset,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical,
        ShapeOutline? Outline,
        ConnectionRef? StartConnection,
        ConnectionRef? EndConnection) : GroupChildModel(LocalRect);

    /// <summary>
    /// グループ内の入れ子グループ。<see cref="ChildOffset"/>/<see cref="ChildExtent"/>は
    /// このグループ自身の孫要素の座標系(親グループとは別の子座標空間)。<see cref="Id"/>は
    /// 接続線の接続先解決(要件10.11)のために保持する。
    /// </summary>
    public sealed record GroupChildGroup(
        uint Id,
        RectPt LocalRect,
        double RotationDegrees,
        PointPt ChildOffset,
        PointPt ChildExtent,
        IReadOnlyList<GroupChildModel> Children) : GroupChildModel(LocalRect);
}
