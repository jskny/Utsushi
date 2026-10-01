using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// 1ページ分の描画命令(背景 → 罫線 → テキスト)を組み立てる。
    /// </summary>
    /// <remarks>
    /// ページに含まれる行/列の並びを受け取り、セルごとの矩形を割り当てたうえで
    /// 結合セルの統合(要件4.3)、罫線・背景(要件4.2)、テキスト配置(要件4.1, 4.4, 2.5)を計算する。
    /// 座標は用紙左上原点で、余白と拡大縮小率を適用済みの最終値として出力する。
    /// </remarks>
    internal sealed class PageCommandBuilder
    {
        private readonly ReportModel _report;
        private readonly SheetModel _sheet;
        private readonly SheetGrid _grid;
        private readonly IFontMetricsProvider _fontMetrics;
        private readonly double _scale;
        private readonly double _originX;
        private readonly double _originY;
        private readonly RectPt? _printableArea;
        private readonly MergedCellIndex? _mergedIndex;

        /// <summary>
        /// 2セルアンカー画像の幅/高さ計算で合算する列/行数の上限。対角セルにセル番地の上限
        /// (最大1,048,576行×16,384列)近くを指定する不正な入力による計算量の増大を防ぐ
        /// (security-reviewer指摘)。自社ロゴ用途でこの上限に達することは想定していない。
        /// </summary>
        private const int MaxSpanCells = 4096;

        /// <summary>画像・図形1つの表示サイズ(pt)の上限。異常に大きいEMU値に対する安全弁。</summary>
        private const double MaxDrawingObjectDimensionPt = 5000.0;

        /// <summary>
        /// <see cref="ConnectorModel.Outline"/>が<c>null</c>の接続線に補う既定の枠線(Excel上は黒い実線1ptで表示される。
        /// design.md参照)。印刷倍率を線幅に掛けるため、Rendering レイヤーではなくここで補う。
        /// </summary>
        private static readonly ShapeOutline DefaultConnectorOutline = new(ArgbColor.Black, 1.0);

        private readonly List<DrawCommand> _fills = new();
        private readonly List<DrawCommand> _borders = new();
        private readonly List<DrawCommand> _texts = new();
        private readonly List<DrawCommand> _drawingObjects = new();
        private readonly HashSet<string> _emittedBorders = new(StringComparer.Ordinal);

        /// <param name="report">帳票。</param>
        /// <param name="grid">印刷対象の格子。</param>
        /// <param name="fontMetrics">フォントメトリクス。</param>
        /// <param name="scale">拡大縮小率。</param>
        /// <param name="margins">余白。本文の左上は(左余白, 上余白)に置く。</param>
        /// <param name="centeringOffset">
        /// 「ページ中央」(<see cref="PageSetupModel.HorizontalCentered"/>/<see cref="PageSetupModel.VerticalCentered"/>)
        /// のために本文全体を平行移動する量(用紙座標、ポイント)。
        /// </param>
        /// <param name="printableArea">
        /// 印字可能領域(余白の内側、用紙座標)。はみ出し表示の文字をこのページの本文の矩形で切り取るときに、
        /// 本文の矩形をさらにこの矩形で切り詰める。null の場合は本文の矩形だけで切り取る。
        /// </param>
        /// <param name="mergedIndex">
        /// セル→結合範囲の索引(<paramref name="grid"/> の行を登録したもの)。null の場合はページごとに作る。
        /// 同じ格子の複数ページで使い回すため、呼び出し側で1度だけ作って渡す。
        /// </param>
        public PageCommandBuilder(
            ReportModel report,
            SheetGrid grid,
            IFontMetricsProvider fontMetrics,
            double scale,
            PageMargins margins,
            PointPt centeringOffset = default,
            RectPt? printableArea = null,
            MergedCellIndex? mergedIndex = null)
        {
            _report = report;
            _sheet = report.Sheet;
            _grid = grid;
            _fontMetrics = fontMetrics;
            _scale = scale;
            _originX = margins.LeftPt + centeringOffset.X;
            _originY = margins.TopPt + centeringOffset.Y;
            _printableArea = printableArea;
            _mergedIndex = mergedIndex;
        }

        /// <summary>はみ出し表示の文字を切り取る矩形(このページの本文の矩形)。<see cref="Build"/> で決まる。</summary>
        private RectPt _pageBodyRect;

        /// <summary>ページに含まれる行・列からの描画命令を生成する。</summary>
        public IReadOnlyList<DrawCommand> Build(IReadOnlyList<int> rows, IReadOnlyList<int> columns)
        {
            var columnOffsets = BuildOffsets(columns, _grid.GetColumnWidthPt);
            var rowOffsets = BuildOffsets(rows, _grid.GetRowHeightPt);

            var columnIndex = BuildIndex(columns);
            var rowIndex = BuildIndex(rows);

            _pageBodyRect = ResolvePageBodyRect(rowOffsets, columnOffsets);

            // セルごとに結合範囲を線形探索しないよう、索引を引く(印刷範囲ごとに作った索引を
            // 呼び出し側で使い回すのが既定。無ければこのページの列について作る)。
            var mergedIndex = _mergedIndex ?? MergedCellIndex.Create(_sheet.MergedRanges, columns);

            // 同一ページ上で同じ結合範囲を二重に描かないための記録。
            var emittedMergedRanges = new HashSet<CellRange>();

            foreach (var row in rows)
            {
                foreach (var column in columns)
                {
                    var address = new CellAddress(row, column);
                    var merged = mergedIndex.Find(address);

                    if (merged is not null)
                    {
                        if (!emittedMergedRanges.Add(merged.Range))
                        {
                            continue;
                        }

                        // 結合範囲がページをまたぐ場合、このページに見えている部分だけを描画する。
                        var mergedRect = TryGetMergedRect(
                            merged.Range, rowIndex, columnIndex, rowOffsets, columnOffsets);
                        if (mergedRect is null)
                        {
                            continue;
                        }

                        var anchorCell = _sheet.GetCell(merged.Anchor);
                        var borders = ResolveMergedBorders(
                            merged.Range, anchorCell?.Style.Borders ?? BorderSet.None, rowIndex, columnIndex);
                        EmitCell(merged.Anchor, anchorCell, mergedRect.Value, borders, merged.Range);
                        continue;
                    }

                    var cell = _sheet.GetCell(address);
                    var rect = CellRect(row, column, rowIndex, columnIndex, rowOffsets, columnOffsets);
                    EmitCell(address, cell, rect);
                }
            }

            EmitDrawingObjects(rowIndex, columnIndex, rowOffsets, columnOffsets);

            var commands = new List<DrawCommand>(_fills.Count + _borders.Count + _texts.Count + _drawingObjects.Count);
            commands.AddRange(_fills);
            commands.AddRange(_borders);
            commands.AddRange(_texts);
            commands.AddRange(_drawingObjects);
            return commands;
        }

        /// <summary>
        /// このページの本文(印刷タイトルを含む)の矩形。印字可能領域が分かっていれば、その内側に切り詰める
        /// (1行/1列だけで印字可能領域を超える場合に、はみ出し表示の文字が余白へ描かれないようにする)。
        /// </summary>
        private RectPt ResolvePageBodyRect(double[] rowOffsets, double[] columnOffsets)
        {
            var body = ToPageRect(0.0, 0.0, columnOffsets[columnOffsets.Length - 1], rowOffsets[rowOffsets.Length - 1]);
            if (_printableArea is not { } area)
            {
                return body;
            }

            var left = Math.Max(body.Left, area.Left);
            var top = Math.Max(body.Top, area.Top);
            return RectPt.FromBounds(
                left,
                top,
                Math.Max(left, Math.Min(body.Right, area.Right)),
                Math.Max(top, Math.Min(body.Bottom, area.Bottom)));
        }

        /// <summary>
        /// 図形・接続線の枠線の太さ(矢印の大きさもこれに比例する)に印刷倍率を掛ける。
        /// 色(透明=線なしを含む)と矢印の種類はそのまま保つ。
        /// </summary>
        private ShapeOutline? ScaleOutline(ShapeOutline? outline) =>
            outline is null ? null : outline with { WidthPt = outline.WidthPt * _scale };

        // ---------------------------------------------------------------------
        // 画像・図形(要件9, 10)
        // ---------------------------------------------------------------------

        /// <summary>
        /// アンカー左上セルがこのページに含まれる画像・図形を、ページ座標に変換して描画命令にする。
        /// <see cref="SheetModel.DrawingObjects"/> の出現順(drawing.xmlの重なり順。要件10.3)を
        /// そのまま維持するため、画像・図形をまとめて1回だけ列挙する。改ページ位置をまたぐ
        /// 画像・図形は、アンカー左上セルが属するページにのみ全体を配置する(design.md「未決事項」の割り切り)。
        /// </summary>
        /// <remarks>
        /// 接続線の接続点解決(要件10.11)のため、この出現順の列挙より先に読み取り専用の予備パス
        /// (<see cref="BuildConnectionTargetTable"/>)でID→ページ矩形のテーブルを構築する。
        /// 予備パスは矩形計算のみでコマンドを一切生成しないため、下の出現順の列挙(z-orderを
        /// 保つ唯一のコマンド生成経路)には手を入れずに、接続線が自分より後に出現する図形を
        /// 参照していても解決できる(design.md参照)。
        /// </remarks>
        private void EmitDrawingObjects(
            IReadOnlyDictionary<int, int> rowIndex,
            IReadOnlyDictionary<int, int> columnIndex,
            double[] rowOffsets,
            double[] columnOffsets)
        {
            var connectionTargets = BuildConnectionTargetTable(rowIndex, columnIndex, rowOffsets, columnOffsets);

            foreach (var drawingObject in _sheet.DrawingObjects)
            {
                if (!TryComputeDrawingObjectRect(drawingObject, rowIndex, columnIndex, rowOffsets, columnOffsets, out var rect))
                {
                    continue;
                }

                switch (drawingObject)
                {
                    case ImageModel image:
                        _drawingObjects.Add(new ImageCommand(rect, image.Data, image.ContentType, image.RotationDegrees));
                        break;
                    case ShapeModel shape:
                        _drawingObjects.Add(BuildShapeCommand(shape, rect));
                        break;
                    case ConnectorModel connector:
                        var (resolvedStart, resolvedEnd) = ResolveConnectorEndpoints(
                            connector.StartConnection, connector.EndConnection, connectionTargets);
                        _drawingObjects.Add(new ConnectorCommand(
                            rect, connector.Preset, connector.RotationDegrees, connector.FlipHorizontal, connector.FlipVertical,
                            ScaleOutline(connector.Outline ?? DefaultConnectorOutline), resolvedStart, resolvedEnd));
                        break;
                    case GroupShapeModel group:
                        var children = BuildGroupChildren(
                            group.Children, rect, group.ChildOffset, group.ChildExtent, connectionTargets,
                            group.FlipHorizontal, group.FlipVertical);
                        _drawingObjects.Add(new GroupCommand(RectCenter(rect), group.RotationDegrees, children));
                        break;
                }
            }
        }

        /// <summary>
        /// 接続点解決(要件10.11)のためのID→ページ矩形の読み取り専用テーブルを構築する。
        /// コマンドは一切生成せず、<see cref="TryComputeDrawingObjectRect"/>/
        /// <see cref="ToGroupChildRect"/>と同じ計算をこの予備パス用に再度行うだけである
        /// (単純な算術のみでコストは無視できる。design.md参照)。接続線自身は接続先として
        /// 参照される対象ではないため、このテーブルには含めない。
        /// </summary>
        /// <remarks>
        /// <c>@id</c>はOOXMLスキーマ上必須かつExcelが重複させないため通常は起こらないが、
        /// 万一同一シート内で重複していた場合、このテーブルは出現順で後から見つかった方の
        /// 図形の座標で上書きする(先勝ちでも後勝ちでもどちらかの図形を選ぶしかなく、
        /// 「解決しない」よりは実害が小さいための割り切り。code-reviewer指摘)。
        /// </remarks>
        private Dictionary<uint, (RectPt Rect, ShapePresetType? Preset, bool FlipHorizontal, bool FlipVertical)> BuildConnectionTargetTable(
            IReadOnlyDictionary<int, int> rowIndex,
            IReadOnlyDictionary<int, int> columnIndex,
            double[] rowOffsets,
            double[] columnOffsets)
        {
            var table = new Dictionary<uint, (RectPt Rect, ShapePresetType? Preset, bool FlipHorizontal, bool FlipVertical)>();

            foreach (var drawingObject in _sheet.DrawingObjects)
            {
                if (!TryComputeDrawingObjectRect(drawingObject, rowIndex, columnIndex, rowOffsets, columnOffsets, out var rect))
                {
                    continue;
                }

                switch (drawingObject)
                {
                    case ImageModel image:
                        table[image.Id] = (rect, null, false, false);
                        break;
                    case ShapeModel shape:
                        table[shape.Id] = (rect, shape.Preset, shape.FlipHorizontal, shape.FlipVertical);
                        break;
                    case GroupShapeModel group:
                        table[group.Id] = (rect, null, group.FlipHorizontal, group.FlipVertical);
                        CollectGroupChildRects(
                            group.Children, rect, group.ChildOffset, group.ChildExtent, table, group.FlipHorizontal, group.FlipVertical);
                        break;
                }
            }

            return table;
        }

        /// <summary>
        /// <see cref="BuildConnectionTargetTable"/>のグループ内再帰部分。<see cref="BuildGroupChildren"/>と
        /// 同じ比例変換(<see cref="ToGroupChildRect"/>)を使うが、コマンドは生成しない。
        /// </summary>
        private static void CollectGroupChildRects(
            IReadOnlyList<GroupChildModel> children,
            RectPt groupRect,
            PointPt childOffset,
            PointPt childExtent,
            Dictionary<uint, (RectPt Rect, ShapePresetType? Preset, bool FlipHorizontal, bool FlipVertical)> table,
            bool flipHorizontal,
            bool flipVertical)
        {
            if (childExtent.X <= 0 || childExtent.Y <= 0)
            {
                return;
            }

            var scaleX = groupRect.Width / childExtent.X;
            var scaleY = groupRect.Height / childExtent.Y;

            foreach (var child in children)
            {
                var childRect = MirrorInGroup(
                    ToGroupChildRect(child.LocalRect, groupRect, childOffset, scaleX, scaleY), groupRect, flipHorizontal, flipVertical);
                if (childRect.IsEmpty)
                {
                    continue;
                }

                switch (child)
                {
                    case GroupChildShape shape:
                        table[shape.Id] = (
                            childRect, shape.Preset, shape.FlipHorizontal ^ flipHorizontal, shape.FlipVertical ^ flipVertical);
                        break;
                    case GroupChildImage image:
                        // 画像そのものは鏡像にしないが、接続点は反転したグループの中での位置に合わせて鏡映する。
                        table[image.Id] = (childRect, null, flipHorizontal, flipVertical);
                        break;
                    case GroupChildGroup nestedGroup:
                        table[nestedGroup.Id] = (
                            childRect, null, nestedGroup.FlipHorizontal ^ flipHorizontal, nestedGroup.FlipVertical ^ flipVertical);
                        CollectGroupChildRects(
                            nestedGroup.Children, childRect, nestedGroup.ChildOffset, nestedGroup.ChildExtent, table,
                            nestedGroup.FlipHorizontal ^ flipHorizontal, nestedGroup.FlipVertical ^ flipVertical);
                        break;
                }
            }
        }

        /// <summary>
        /// 接続線の始点/終点の接続先(<see cref="ConnectionRef"/>)を、
        /// <paramref name="connectionTargets"/>を使って実際の座標へ解決する。参照先が
        /// このページに無い、IDが実在しない、または接続先の指定自体が無い場合は<c>null</c>のままとし、
        /// Renderingレイヤーが要件10.9の既定動作(アンカー矩形+反転)にフォールバックする。
        /// </summary>
        private static (PointPt? Start, PointPt? End) ResolveConnectorEndpoints(
            ConnectionRef? startConnection,
            ConnectionRef? endConnection,
            IReadOnlyDictionary<uint, (RectPt Rect, ShapePresetType? Preset, bool FlipHorizontal, bool FlipVertical)> connectionTargets) =>
            (ResolveConnectionPoint(startConnection, connectionTargets), ResolveConnectionPoint(endConnection, connectionTargets));

        private static PointPt? ResolveConnectionPoint(
            ConnectionRef? connection, IReadOnlyDictionary<uint, (RectPt Rect, ShapePresetType? Preset, bool FlipHorizontal, bool FlipVertical)> connectionTargets)
        {
            if (connection is not { } reference || !connectionTargets.TryGetValue(reference.ShapeId, out var target))
            {
                return null;
            }

            return ConnectionSiteResolver.Resolve(
                target.Rect, target.Preset, reference.SiteIndex, target.FlipHorizontal, target.FlipVertical);
        }

        /// <summary>
        /// グループの子座標空間(<paramref name="childOffset"/>/<paramref name="childExtent"/>)から
        /// <paramref name="groupRect"/>(グループ自身のページ矩形)への比例変換(平行移動+拡大縮小、
        /// 非一様倍率を許容する)で各子要素をページ座標へ変換する(要件10.10)。
        /// 入れ子の<see cref="GroupChildGroup"/>は自身のページ矩形を新たな基準として再帰的に適用する。
        /// </summary>
        private List<DrawCommand> BuildGroupChildren(
            IReadOnlyList<GroupChildModel> children,
            RectPt groupRect,
            PointPt childOffset,
            PointPt childExtent,
            IReadOnlyDictionary<uint, (RectPt Rect, ShapePresetType? Preset, bool FlipHorizontal, bool FlipVertical)> connectionTargets,
            bool flipHorizontal,
            bool flipVertical)
        {
            var result = new List<DrawCommand>(children.Count);

            // グループの反転(要件10.17)は、子要素の配置矩形をグループの矩形内で鏡映し、子要素自身の反転を
            // 切り替えることで畳み込む。片方向だけの鏡映は回転の向きも逆にする(時計回り θ の鏡像は -θ)。
            var mirrorsRotation = flipHorizontal ^ flipVertical;
            double Rotation(double degrees) => mirrorsRotation ? -degrees : degrees;

            // 子座標空間の大きさが0以下では比例変換できないため、このグループの子要素は
            // 何も描画しない(壊れたジオメトリに対する安全弁)。
            if (childExtent.X <= 0 || childExtent.Y <= 0)
            {
                return result;
            }

            var scaleX = groupRect.Width / childExtent.X;
            var scaleY = groupRect.Height / childExtent.Y;

            foreach (var child in children)
            {
                var childRect = MirrorInGroup(
                    ToGroupChildRect(child.LocalRect, groupRect, childOffset, scaleX, scaleY), groupRect, flipHorizontal, flipVertical);
                if (childRect.IsEmpty)
                {
                    continue;
                }

                switch (child)
                {
                    case GroupChildShape shape:
                        // テキスト余白・フォントサイズにグループ自身のリサイズ比率(scaleX/scaleY)を
                        // 反映するため、幾何平均を「等方的な」代表スケールとして渡す
                        // (layout-fidelity-reviewer指摘: 以前は印刷拡大率(_scale)のみを見ており、
                        // グループが大きく縮小されている場合に余白がシェイプ本体ほど縮まらず、
                        // 縮小率次第ではテキストが矩形からはみ出す/消える境界に達しうる)。
                        var groupScale = Math.Sqrt(Math.Abs(scaleX * scaleY));
                        result.Add(BuildGroupChildShapeCommand(
                            shape, childRect, groupScale, Rotation(shape.RotationDegrees),
                            shape.FlipHorizontal ^ flipHorizontal, shape.FlipVertical ^ flipVertical));
                        break;
                    case GroupChildImage image:
                        // 画像の中身の鏡像化には対応しない(配置と回転の向きだけを反映する)。
                        result.Add(new ImageCommand(childRect, image.Data, image.ContentType, Rotation(image.RotationDegrees)));
                        break;
                    case GroupChildConnector connector:
                        var (resolvedStart, resolvedEnd) = ResolveConnectorEndpoints(
                            connector.StartConnection, connector.EndConnection, connectionTargets);
                        result.Add(new ConnectorCommand(
                            childRect, connector.Preset, Rotation(connector.RotationDegrees),
                            connector.FlipHorizontal ^ flipHorizontal, connector.FlipVertical ^ flipVertical,
                            ScaleOutline(connector.Outline ?? DefaultConnectorOutline), resolvedStart, resolvedEnd));
                        break;
                    case GroupChildGroup nestedGroup:
                        var nestedChildren = BuildGroupChildren(
                            nestedGroup.Children, childRect, nestedGroup.ChildOffset, nestedGroup.ChildExtent, connectionTargets,
                            nestedGroup.FlipHorizontal ^ flipHorizontal, nestedGroup.FlipVertical ^ flipVertical);
                        result.Add(new GroupCommand(RectCenter(childRect), Rotation(nestedGroup.RotationDegrees), nestedChildren));
                        break;
                }
            }

            return result;
        }

        /// <summary>
        /// グループの子座標空間上の矩形(<paramref name="localRect"/>)を、
        /// <paramref name="groupRect"/>を基準とした比例変換でページ座標の矩形へ変換する。
        /// </summary>
        private static RectPt ToGroupChildRect(
            RectPt localRect, RectPt groupRect, PointPt childOffset, double scaleX, double scaleY)
        {
            var left = groupRect.Left + ((localRect.Left - childOffset.X) * scaleX);
            var top = groupRect.Top + ((localRect.Top - childOffset.Y) * scaleY);
            var width = localRect.Width * scaleX;
            var height = localRect.Height * scaleY;

            // 異常に大きいEMU値による過大な矩形を防ぐ(画像・図形と同じ安全弁)。
            width = Math.Min(width, MaxDrawingObjectDimensionPt);
            height = Math.Min(height, MaxDrawingObjectDimensionPt);

            return width <= 0 || height <= 0
                ? default
                : RectPt.FromBounds(left, top, left + width, top + height);
        }

        /// <summary>
        /// グループ内図形の描画命令を組み立てる(要件10.10)。テキスト折り返しはトップレベルの図形と
        /// 同じロジックを再利用するが、<paramref name="groupScale"/>でグループ自身のリサイズ比率を
        /// 追加で反映する(トップレベルの図形は既定の1.0のまま、<see cref="BuildShapeCommand"/>参照)。
        /// </summary>
        private ShapeCommand BuildGroupChildShapeCommand(
            GroupChildShape shape, RectPt rect, double groupScale, double rotationDegrees, bool flipHorizontal, bool flipVertical)
        {
            var textLines = shape.Text is { } text
                ? BuildShapeTextLines(text, rect, groupScale)
                : Array.Empty<ShapeTextLine>();

            return new ShapeCommand(
                rect, shape.Preset, shape.AdjustmentValues, rotationDegrees, shape.Fill, ScaleOutline(shape.Outline), textLines,
                flipHorizontal, flipVertical);
        }

        /// <summary>
        /// 反転したグループ(要件10.17)の子要素の矩形を、グループの矩形の中心を軸に鏡映する。
        /// </summary>
        private static RectPt MirrorInGroup(RectPt childRect, RectPt groupRect, bool flipHorizontal, bool flipVertical)
        {
            if (childRect.IsEmpty || (!flipHorizontal && !flipVertical))
            {
                return childRect;
            }

            var left = flipHorizontal ? groupRect.Left + groupRect.Right - childRect.Right : childRect.Left;
            var top = flipVertical ? groupRect.Top + groupRect.Bottom - childRect.Bottom : childRect.Top;
            return RectPt.FromBounds(left, top, left + childRect.Width, top + childRect.Height);
        }

        private static PointPt RectCenter(RectPt rect) => new(rect.Left + (rect.Width / 2.0), rect.Top + (rect.Height / 2.0));

        /// <summary>
        /// 画像・図形共通のアンカー解決(要件9.1, 9.2, 10.1, 10.2)。アンカー左上セルがこのページに
        /// 含まれない場合は <c>false</c> を返す。
        /// </summary>
        private bool TryComputeDrawingObjectRect(
            DrawingObjectModel drawingObject,
            IReadOnlyDictionary<int, int> rowIndex,
            IReadOnlyDictionary<int, int> columnIndex,
            double[] rowOffsets,
            double[] columnOffsets,
            out RectPt rect)
        {
            rect = default;

            if (!rowIndex.TryGetValue(drawingObject.AnchorCell.Row, out var r)
                || !columnIndex.TryGetValue(drawingObject.AnchorCell.Column, out var c))
            {
                return false;
            }

            var left = columnOffsets[c] + drawingObject.AnchorOffset.X;
            var top = rowOffsets[r] + drawingObject.AnchorOffset.Y;

            var (widthPt, heightPt) = drawingObject.Extent switch
            {
                FixedAnchorExtent fixedExtent => (fixedExtent.WidthPt, fixedExtent.HeightPt),
                CellSpanAnchorExtent span => (
                    SpanWidthPt(drawingObject.AnchorCell.Column, drawingObject.AnchorOffset.X, span.ToCell.Column, span.ToOffset.X),
                    SpanHeightPt(drawingObject.AnchorCell.Row, drawingObject.AnchorOffset.Y, span.ToCell.Row, span.ToOffset.Y)),
                _ => (0.0, 0.0),
            };

            if (widthPt <= 0 || heightPt <= 0)
            {
                return false;
            }

            rect = ToPageRect(left, top, left + widthPt, top + heightPt);

            // 異常に大きいEMU値(またはその合算)による過大な矩形を防ぐ(security-reviewer指摘)。
            // 上限は印刷拡大率(_scale)適用後の最終的な表示サイズに対して適用する
            // (グループ内子要素のToGroupChildRectと適用点を揃える。layout-fidelity-reviewer指摘:
            // 以前は_scale適用前のwidthPt/heightPtに上限を適用しており、印刷拡大率が100%を超える
            // 帳票では上限の実効値がグループ内子要素と食い違っていた)。
            var cappedWidth = Math.Min(rect.Width, MaxDrawingObjectDimensionPt);
            var cappedHeight = Math.Min(rect.Height, MaxDrawingObjectDimensionPt);
            rect = RectPt.FromBounds(rect.Left, rect.Top, rect.Left + cappedWidth, rect.Top + cappedHeight);

            return !rect.IsEmpty;
        }

        /// <summary>図形の描画命令を組み立てる(要件10)。テキストの折り返し・配置はここで確定させ、回転前のローカル座標で保持する(回転はRenderingレイヤーが適用する)。</summary>
        private ShapeCommand BuildShapeCommand(ShapeModel shape, RectPt rect)
        {
            var textLines = shape.Text is { } text
                ? BuildShapeTextLines(text, rect)
                : Array.Empty<ShapeTextLine>();

            return new ShapeCommand(
                rect, shape.Preset, shape.AdjustmentValues, shape.RotationDegrees, shape.Fill, ScaleOutline(shape.Outline), textLines,
                shape.FlipHorizontal, shape.FlipVertical);
        }

        /// <summary>
        /// 図形内テキストを矩形幅で折り返し、水平/垂直配置に基づく各行のローカル座標を確定させる(要件10.4)。
        /// </summary>
        /// <param name="groupScale">
        /// グループ内図形の場合の、グループ自身のリサイズ比率(<see cref="BuildGroupChildren"/>が
        /// scaleX/scaleYの幾何平均として算出)。トップレベルの図形は1.0(<see cref="BuildShapeCommand"/>)。
        /// </param>
        private IReadOnlyList<ShapeTextLine> BuildShapeTextLines(ShapeTextBody text, RectPt rect, double groupScale = 1.0)
        {
            // 拡大縮小率はセル内テキスト(EmitText)と同様、余白・フォントサイズの両方に適用する
            // (layout-fidelity-reviewer指摘: 図形の矩形自体はToPageRectで_scaleが掛かるのに、
            // 内側のテキストが原寸のままだと、印刷倍率を持つ帳票でテキストが矩形からはみ出す)。
            // groupScaleは、グループ内図形の矩形自体がグループのリサイズ比率で既に縮小/拡大されている
            // ことに合わせて余白・フォントサイズも追従させるための追加の係数
            // (layout-fidelity-reviewer指摘: グループが大きく縮小されている場合、以前は余白が
            // _scale分しか縮まらずシェイプ本体ほど縮小されないため、縮小率次第でcontentRectが
            // 0以下になりテキストが消える境界に達しうる)。
            // 余白は a:bodyPr の lIns/tIns/rIns/bIns(指定が無ければ DrawingML の既定値 左右7.2pt・上下3.6pt)。
            var insets = text.Insets ?? ShapeTextInsets.Default;
            var insetScale = _scale * groupScale;
            var contentRect = RectPt.FromBounds(
                rect.Left + (insets.LeftPt * insetScale),
                rect.Top + (insets.TopPt * insetScale),
                rect.Right - (insets.RightPt * insetScale),
                rect.Bottom - (insets.BottomPt * insetScale));

            if (contentRect.Width <= 0 || contentRect.Height <= 0)
            {
                return Array.Empty<ShapeTextLine>();
            }

            var wrapped = WrapShapeText(text, contentRect.Width, groupScale);
            if (wrapped.Count == 0)
            {
                return Array.Empty<ShapeTextLine>();
            }

            var firstMetrics = _fontMetrics.GetMetrics(wrapped[0].Font);
            var totalHeight = wrapped.Sum(line => _fontMetrics.GetMetrics(line.Font).LineSpacingPt);
            var y = ResolveFirstBaselineY(text.VAlign, contentRect, firstMetrics, totalHeight);

            var lines = new List<ShapeTextLine>(wrapped.Count);
            foreach (var (lineText, hAlign, font) in wrapped)
            {
                var metrics = _fontMetrics.GetMetrics(font);
                var (x, anchor) = ResolveTextOrigin(hAlign, contentRect);
                lines.Add(new ShapeTextLine(new PointPt(x, y), lineText, font, anchor));
                y += metrics.LineSpacingPt;
            }

            return lines;
        }

        /// <summary>
        /// 図形内テキストの各段落を、その段落の先頭ランのフォントで折り返す(要件10.4)。
        /// 段落内で複数ランが異なるフォントを持つ場合でも、折り返し計算は先頭ランのフォントで代表させる
        /// (図形は注記・吹き出し用途を想定した近似実装であり、セル内テキストほど厳密な混在対応はしない)。
        /// </summary>
        /// <remarks>
        /// 段落内の改行(<c>a:br</c>。ランの中の改行文字)は行の区切りとして扱い(<see cref="WrapLines"/>)、
        /// ランを持たない空の段落は1行分の空行にする。空の段落の行の高さは、直前(無ければ直後)の
        /// ランを持つ段落の先頭ランのフォントで代表させる。
        /// </remarks>
        private List<(string Text, HorizontalAlignment HAlign, FontStyle Font)> WrapShapeText(
            ShapeTextBody text, double availableWidthPt, double groupScale = 1.0)
        {
            var lines = new List<(string, HorizontalAlignment, FontStyle)>();
            var paragraphs = text.Paragraphs;
            for (var p = 0; p < paragraphs.Count; p++)
            {
                var paragraph = paragraphs[p];

                // 拡大縮小率はフォントサイズにも適用する(セル内テキストのEmitTextと同様。
                // 座標だけを縮めると文字が矩形に収まらなくなるため)。groupScaleは
                // グループ内図形の場合の追加のリサイズ比率(BuildShapeTextLines参照)。
                var baseFont = RepresentativeFont(paragraphs, p);
                var scaledFont = baseFont with { SizePt = baseFont.SizePt * _scale * groupScale };

                if (paragraph.Runs.Count == 0)
                {
                    lines.Add((string.Empty, paragraph.HAlign, scaledFont));
                    continue;
                }

                var paragraphText = string.Concat(paragraph.Runs.Select(run => run.Text));
                foreach (var line in WrapLines(scaledFont, paragraphText, availableWidthPt))
                {
                    lines.Add((line, paragraph.HAlign, scaledFont));
                }
            }

            return lines;
        }

        /// <summary>
        /// 段落 <paramref name="index"/> の折り返し・行の高さに使うフォント。ランを持つ段落はその先頭ランのフォント、
        /// 空の段落は直前(無ければ直後)のランを持つ段落の先頭ランのフォント。どこにも無ければ既定のフォント。
        /// </summary>
        private static FontStyle RepresentativeFont(IReadOnlyList<ShapeTextParagraph> paragraphs, int index)
        {
            for (var i = index; i >= 0; i--)
            {
                if (paragraphs[i].Runs.Count > 0)
                {
                    return paragraphs[i].Runs[0].Font;
                }
            }

            for (var i = index + 1; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].Runs.Count > 0)
                {
                    return paragraphs[i].Runs[0].Font;
                }
            }

            return FontStyle.Default;
        }

        /// <summary>
        /// 2セルアンカー(<see cref="CellSpanAnchorExtent"/>)の幅を求める。列幅の合計は
        /// <see cref="_grid"/>(印刷範囲にクリップされた格子)ではなく <see cref="_sheet"/> から
        /// 直接取得する。<see cref="_grid"/> は印刷範囲外の列を「幅0」として保持しないため、
        /// 対角セルが印刷範囲のすぐ外にあるだけで画像が実際より小さく計算されてしまう
        /// (layout-fidelity-reviewer指摘の不具合)。Excel自体は印刷範囲の設定に関わらず
        /// 実際の列幅で画像サイズを決めるため、これに合わせる(非表示列は0として扱う点は
        /// 結合セル等の既存ロジックと同様)。列数には上限を設け、対角セルにセル番地の上限
        /// (最大1,048,576行×16,384列)近くを指定する不正な入力による計算量の増大を防ぐ
        /// (security-reviewer指摘)。
        /// </summary>
        private double SpanWidthPt(int fromColumn, double fromOffsetPt, int toColumn, double toOffsetPt)
        {
            var width = toOffsetPt - fromOffsetPt;
            var lastColumn = Math.Min(toColumn, fromColumn + MaxSpanCells);
            for (var column = fromColumn; column < lastColumn; column++)
            {
                width += RawColumnWidthPt(column);
            }

            return Math.Max(0.0, width);
        }

        /// <summary>2セルアンカーの高さを求める。<see cref="SpanWidthPt"/>と同様の考え方。</summary>
        private double SpanHeightPt(int fromRow, double fromOffsetPt, int toRow, double toOffsetPt)
        {
            var height = toOffsetPt - fromOffsetPt;
            var lastRow = Math.Min(toRow, fromRow + MaxSpanCells);
            for (var row = fromRow; row < lastRow; row++)
            {
                height += RawRowHeightPt(row);
            }

            return Math.Max(0.0, height);
        }

        /// <summary>印刷範囲によらない、シート上の実際の列幅(pt)。非表示列は0。</summary>
        private double RawColumnWidthPt(int column) =>
            SheetGrid.PrintedColumnWidthPt(_sheet, column, _report.Definition.MaxDigitWidthPx);

        /// <summary>印刷範囲によらない、シート上の実際の行高(pt)。非表示行は0。</summary>
        private double RawRowHeightPt(int row) => SheetGrid.PrintedRowHeightPt(_sheet, row);

        private void EmitCell(
            CellAddress address, CellModel? cell, RectPt rect, BorderSet? bordersOverride = null, CellRange? mergedRange = null)
        {
            var style = cell?.Style ?? CellStyle.Default;
            var text = cell?.DisplayValue;

            // 差し込み値は、矩形が空(行高・列幅が0)で何も描かれない場合も含めて検証する(要件2.13, 2.14)。
            if (!string.IsNullOrEmpty(text) && _report.SubstitutedCells.Contains(address))
            {
                EnsureSubstitutedTextFits(address, style, text!, rect, mergedRange);
            }

            if (rect.IsEmpty)
            {
                return;
            }

            if (!style.BackgroundColor.IsTransparent)
            {
                _fills.Add(new FillRectCommand(rect, style.BackgroundColor));
            }

            EmitBorders(rect, bordersOverride ?? style.Borders);

            if (!string.IsNullOrEmpty(text))
            {
                EmitText(address, cell!, style, rect, text!, mergedRange);
            }
        }

        /// <summary>
        /// 結合範囲の外周4辺の罫線を、範囲を構成する各セルから辺ごとに合成する。
        /// </summary>
        /// <remarks>
        /// <para>
        /// Excelで結合範囲に外枠/格子を設定すると、右端の罫線は右端列の各セルのRight、
        /// 下端の罫線は下端行の各セルのBottomに保持される。アンカー(左上)セルの罫線だけを
        /// 見ると、アンカー以外にしか設定されていない罫線(右端・下端など)が欠落する。
        /// </para>
        /// <para>
        /// 結合範囲が改ページをまたぐ場合、このページに見えているのは範囲の一部だけである。
        /// 範囲の本当の上端/下端/左端/右端がこのページに含まれていない辺は、
        /// 単なる改ページの切れ目でしかないため罫線を出さない(<paramref name="rowIndex"/>/
        /// <paramref name="columnIndex"/> で判定する)。
        /// </para>
        /// 斜め罫線はExcel上もアンカーセルのものしか意味を持たないため、そのまま使う。
        /// </remarks>
        private BorderSet ResolveMergedBorders(
            CellRange range,
            BorderSet anchorBorders,
            IReadOnlyDictionary<int, int> rowIndex,
            IReadOnlyDictionary<int, int> columnIndex)
        {
            var left = columnIndex.ContainsKey(range.FirstColumn)
                ? ResolveColumnEdge(range.FirstColumn, range.FirstRow, range.LastRow, b => b.Left)
                : BorderEdge.None;
            var right = columnIndex.ContainsKey(range.LastColumn)
                ? ResolveColumnEdge(range.LastColumn, range.FirstRow, range.LastRow, b => b.Right)
                : BorderEdge.None;
            var top = rowIndex.ContainsKey(range.FirstRow)
                ? ResolveRowEdge(range.FirstRow, range.FirstColumn, range.LastColumn, b => b.Top)
                : BorderEdge.None;
            var bottom = rowIndex.ContainsKey(range.LastRow)
                ? ResolveRowEdge(range.LastRow, range.FirstColumn, range.LastColumn, b => b.Bottom)
                : BorderEdge.None;

            return new BorderSet(left, right, top, bottom, anchorBorders.DiagonalDown, anchorBorders.DiagonalUp);
        }

        /// <summary>
        /// 結合範囲の列方向の枠線を、範囲内の各セルを走査して探す。走査するセル数には
        /// <see cref="SpanWidthPt"/>等と同じ<see cref="MaxSpanCells"/>の上限を設け、
        /// 結合範囲の対角セルにセル番地の上限近く(最大1,048,576行)を指定する不正な入力
        /// (`mergeCell`はParsingレイヤーでサイズ上限を設けていない)による計算量の増大を防ぐ
        /// (code-reviewer指摘)。
        /// </summary>
        private BorderEdge ResolveColumnEdge(int column, int firstRow, int lastRow, Func<BorderSet, BorderEdge> selector)
        {
            var boundedLastRow = Math.Min(lastRow, firstRow + MaxSpanCells);
            for (var row = firstRow; row <= boundedLastRow; row++)
            {
                var edge = selector(_sheet.GetCell(new CellAddress(row, column))?.Style.Borders ?? BorderSet.None);
                if (edge.IsVisible)
                {
                    return edge;
                }
            }

            return BorderEdge.None;
        }

        /// <summary>結合範囲の行方向の枠線を探す。<see cref="ResolveColumnEdge"/>と同様の考え方。</summary>
        private BorderEdge ResolveRowEdge(int row, int firstColumn, int lastColumn, Func<BorderSet, BorderEdge> selector)
        {
            var boundedLastColumn = Math.Min(lastColumn, firstColumn + MaxSpanCells);
            for (var column = firstColumn; column <= boundedLastColumn; column++)
            {
                var edge = selector(_sheet.GetCell(new CellAddress(row, column))?.Style.Borders ?? BorderSet.None);
                if (edge.IsVisible)
                {
                    return edge;
                }
            }

            return BorderEdge.None;
        }

        // ---------------------------------------------------------------------
        // 罫線(要件4.2)
        // ---------------------------------------------------------------------

        private void EmitBorders(RectPt rect, BorderSet borders)
        {
            if (!borders.HasAnyVisibleEdge)
            {
                return;
            }

            EmitEdge(borders.Top, new PointPt(rect.Left, rect.Top), new PointPt(rect.Right, rect.Top), horizontal: true);
            EmitEdge(borders.Bottom, new PointPt(rect.Left, rect.Bottom), new PointPt(rect.Right, rect.Bottom), horizontal: true);
            EmitEdge(borders.Left, new PointPt(rect.Left, rect.Top), new PointPt(rect.Left, rect.Bottom), horizontal: false);
            EmitEdge(borders.Right, new PointPt(rect.Right, rect.Top), new PointPt(rect.Right, rect.Bottom), horizontal: false);

            if (borders.DiagonalDown.IsVisible)
            {
                EmitEdge(
                    borders.DiagonalDown,
                    new PointPt(rect.Left, rect.Top),
                    new PointPt(rect.Right, rect.Bottom),
                    horizontal: false);
            }

            if (borders.DiagonalUp.IsVisible)
            {
                EmitEdge(
                    borders.DiagonalUp,
                    new PointPt(rect.Left, rect.Bottom),
                    new PointPt(rect.Right, rect.Top),
                    horizontal: false);
            }
        }

        private void EmitEdge(BorderEdge edge, PointPt from, PointPt to, bool horizontal)
        {
            if (!edge.IsVisible)
            {
                return;
            }

            var width = BorderMetrics.GetWidthPt(edge.Style) * _scale;
            var dash = BorderMetrics.GetDash(edge.Style);

            if (BorderMetrics.IsDouble(edge.Style))
            {
                // 二重線は細線2本で表現する。罫線が属するセル境界の内外へ等距離に振り分ける。
                var offset = BorderMetrics.DoubleLineCenterSpacingPt * _scale / 2.0;
                var (dx, dy) = horizontal ? (0.0, offset) : (offset, 0.0);
                AddLine(Offset(from, -dx, -dy), Offset(to, -dx, -dy), edge.Color, width, dash);
                AddLine(Offset(from, dx, dy), Offset(to, dx, dy), edge.Color, width, dash);
                return;
            }

            AddLine(from, to, edge.Color, width, dash);
        }

        private static PointPt Offset(PointPt point, double dx, double dy) => new(point.X + dx, point.Y + dy);

        /// <summary>
        /// 罫線を追加する。隣接セルが同じ境界に同一の罫線を持つ場合、描画命令が重複するため取り除く。
        /// </summary>
        private void AddLine(PointPt from, PointPt to, ArgbColor color, double widthPt, LineDashStyle dash)
        {
            var key = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                "{0:0.###},{1:0.###},{2:0.###},{3:0.###},{4},{5:0.###},{6}",
                from.X, from.Y, to.X, to.Y, color, widthPt, dash);

            if (!_emittedBorders.Add(key))
            {
                return;
            }

            _borders.Add(new LineCommand(from, to, color, widthPt, dash));
        }

        // ---------------------------------------------------------------------
        // テキスト(要件4.1, 4.4, 2.5)
        // ---------------------------------------------------------------------

        private void EmitText(
            CellAddress address, CellModel cell, CellStyle style, RectPt rect, string text, CellRange? mergedRange)
        {
            var maxDigitWidthPx = _report.Definition.MaxDigitWidthPx;
            var paddingPt = ExcelUnitConverter.CellPaddingPoints * _scale;
            var indentPt = ExcelUnitConverter.IndentWidthToPoints(style.Indent, maxDigitWidthPx) * _scale;

            var contentRect = RectPt.FromBounds(
                rect.Left + paddingPt + indentPt,
                rect.Top,
                rect.Right - paddingPt,
                rect.Bottom);

            if (contentRect.Width <= 0)
            {
                return;
            }

            var hAlign = ResolveHorizontalAlignment(style.HAlign, cell.ValueKind);
            var overflow = ResolveOverflow(address, style);

            // 拡大縮小率はフォントサイズにも適用する(座標だけを縮めると文字が収まらなくなるため)。
            var scaledFont = style.Font with { SizePt = style.Font.SizePt * _scale };

            // 折り返し表示でないセルでは、Excelは改行を表示せず1行につなげて表示する(要件4.7)。
            // 改行文字をそのまま描画すると、フォントによって豆腐や空白になる。
            var lines = overflow == OverflowBehavior.Wrap
                ? WrapLines(scaledFont, text, contentRect.Width)
                : new List<string> { RemoveLineBreaks(text) };

            if (overflow == OverflowBehavior.Shrink && lines.Count == 1)
            {
                scaledFont = ShrinkToFit(scaledFont, lines[0], contentRect.Width);
            }

            var metrics = _fontMetrics.GetMetrics(scaledFont);
            var totalHeight = metrics.LineSpacingPt * lines.Count;
            var firstBaselineY = ResolveFirstBaselineY(style.VAlign, rect, metrics, totalHeight);

            // はみ出し表示でも、結合範囲の文字は結合範囲の外へ出さない(Excel は結合セルの文字を隣のセルへ
            // はみ出させない)。結合範囲でないセルのはみ出し表示は隣のセルへ描くが、このページの本文の矩形
            // (余白の内側)の外へは描かない。
            var clipsToCell = overflow is OverflowBehavior.Clip or OverflowBehavior.Wrap or OverflowBehavior.Shrink
                || mergedRange is not null;
            var clipRect = clipsToCell ? rect : _pageBodyRect;

            for (var i = 0; i < lines.Count; i++)
            {
                var baselineY = firstBaselineY + (metrics.LineSpacingPt * i);
                var (x, anchor) = ResolveTextOrigin(hAlign, contentRect);
                _texts.Add(new TextCommand(new PointPt(x, baselineY), lines[i], scaledFont, anchor, clipRect));
            }
        }

        /// <summary>
        /// 差し込み値がセル内に表示されることを確かめる(要件2.13, 2.14)。
        /// </summary>
        /// <remarks>
        /// <para>
        /// 判定はページ上に見えている矩形ではなく、セル(結合範囲なら範囲全体)のシート上の本来の大きさで行う。
        /// 結合範囲が改ページ・印刷範囲の端・非表示行にかかって一部しか見えないページでも、判定を省かずに済むようにするため。
        /// </para>
        /// <para>
        /// 折り返し表示では、各行の字面(アセント+ディセント)の4分の1を超えてセルの外に出る行を
        /// 「表示されない」とみなす。フォントの行送りはExcelの行高よりわずかに大きいことが多く、
        /// 字面全体が収まることを条件にすると、Excel上で収まっている住所まで誤検出するため。
        /// </para>
        /// </remarks>
        private void EnsureSubstitutedTextFits(
            CellAddress address, CellStyle style, string text, RectPt visibleRect, CellRange? mergedRange)
        {
            var (widthPt, heightPt) = mergedRange is { } range
                ? (SumBounded(range.FirstColumn, range.LastColumn, RawColumnWidthPt) * _scale,
                   SumBounded(range.FirstRow, range.LastRow, RawRowHeightPt) * _scale)
                : (visibleRect.Width, visibleRect.Height);

            var paddingPt = ExcelUnitConverter.CellPaddingPoints * _scale;
            var indentPt = ExcelUnitConverter.IndentWidthToPoints(style.Indent, _report.Definition.MaxDigitWidthPx) * _scale;
            var contentWidthPt = widthPt - (paddingPt * 2) - indentPt;

            if (contentWidthPt <= 0 || heightPt <= 0)
            {
                throw new LayoutComputationException(
                    $"セル {address} に値を差し込みましたが、セルの幅または高さが無いため、PDFに出力されません。"
                    + "テンプレートの列幅・行の高さを見直してください。",
                    _report.Definition.ReportCode,
                    _sheet.Name,
                    address);
            }

            if (ResolveOverflow(address, style) != OverflowBehavior.Wrap)
            {
                return;
            }

            var font = style.Font with { SizePt = style.Font.SizePt * _scale };
            var lineCount = WrapLines(font, text, contentWidthPt).Count;
            var metrics = _fontMetrics.GetMetrics(font);
            var cellRect = new RectPt(0, 0, widthPt, heightPt);
            var firstBaselineY = ResolveFirstBaselineY(style.VAlign, cellRect, metrics, metrics.LineSpacingPt * lineCount);
            var tolerancePt = (metrics.AscentPt + metrics.DescentPt) / 4.0;

            for (var i = 0; i < lineCount; i++)
            {
                var baselineY = firstBaselineY + (metrics.LineSpacingPt * i);
                var glyphTop = baselineY - metrics.AscentPt;
                var glyphBottom = baselineY + metrics.DescentPt;
                if (glyphTop >= cellRect.Top - tolerancePt && glyphBottom <= cellRect.Bottom + tolerancePt)
                {
                    continue;
                }

                throw new LayoutComputationException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "セル {0} に差し込んだ値は折り返すと {1} 行になり、セルの高さ({2:0.#}pt)に収まらないため、"
                            + "一部の行がPDFに出力されません。テンプレートの行の高さを広げるか、"
                            + "差し込む文字数・行数を減らしてください。",
                        address,
                        lineCount,
                        heightPt / _scale),
                    _report.Definition.ReportCode,
                    _sheet.Name,
                    address);
            }
        }

        /// <summary>行/列番号の範囲の大きさを合計する。走査数には<see cref="MaxSpanCells"/>の上限を設ける。</summary>
        private static double SumBounded(int first, int last, Func<int, double> sizeOf)
        {
            var total = 0.0;
            var boundedLast = Math.Min(last, first + MaxSpanCells);
            for (var i = first; i <= boundedLast; i++)
            {
                total += sizeOf(i);
            }

            return total;
        }

        /// <summary>折り返し表示でないセルのために改行文字を取り除く(要件4.7)。</summary>
        private static string RemoveLineBreaks(string text) =>
            text.IndexOf('\n') < 0 && text.IndexOf('\r') < 0
                ? text
                : text.Replace("\r", string.Empty).Replace("\n", string.Empty);

        /// <summary>
        /// 水平配置を決定する。<see cref="HorizontalAlignment.General"/> は値の型で既定が変わる。
        /// </summary>
        private static HorizontalAlignment ResolveHorizontalAlignment(HorizontalAlignment align, CellValueKind kind)
        {
            if (align != HorizontalAlignment.General)
            {
                return align;
            }

            return kind switch
            {
                CellValueKind.Number => HorizontalAlignment.Right,
                CellValueKind.Boolean => HorizontalAlignment.Center,
                CellValueKind.Error => HorizontalAlignment.Center,
                _ => HorizontalAlignment.Left,
            };
        }

        private static (double X, TextAnchor Anchor) ResolveTextOrigin(HorizontalAlignment align, RectPt contentRect) =>
            align switch
            {
                HorizontalAlignment.Right => (contentRect.Right, TextAnchor.Right),
                HorizontalAlignment.Center or HorizontalAlignment.CenterContinuous =>
                    (contentRect.Left + (contentRect.Width / 2.0), TextAnchor.Center),
                _ => (contentRect.Left, TextAnchor.Left),
            };

        private static double ResolveFirstBaselineY(
            VerticalAlignment align, RectPt rect, FontMetrics metrics, double totalHeight)
        {
            // Excel は行の下端にテキストの下端を合わせるのが既定。
            return align switch
            {
                VerticalAlignment.Top => rect.Top + metrics.AscentPt,
                VerticalAlignment.Center or VerticalAlignment.Distributed or VerticalAlignment.Justify =>
                    rect.Top + ((rect.Height - totalHeight) / 2.0) + metrics.AscentPt,
                _ => rect.Bottom - totalHeight + metrics.AscentPt,
            };
        }

        /// <summary>
        /// はみ出し時の挙動を決定する。帳票定義で置換対象に指定された挙動が、セル書式より優先される。
        /// </summary>
        private OverflowBehavior ResolveOverflow(CellAddress address, CellStyle style)
        {
            if (_report.GetOverflowBehavior(address) is { } fromDefinition)
            {
                return fromDefinition;
            }

            if (style.WrapText)
            {
                return OverflowBehavior.Wrap;
            }

            return style.ShrinkToFit ? OverflowBehavior.Shrink : OverflowBehavior.Overflow;
        }

        /// <summary>
        /// セル幅に収まるようフォントサイズを縮める(Excel の「縮小して全体を表示する」相当)。
        /// </summary>
        private FontStyle ShrinkToFit(FontStyle font, string text, double availableWidthPt)
        {
            var width = _fontMetrics.MeasureTextWidth(font, text);
            if (width <= availableWidthPt || width <= 0)
            {
                return font;
            }

            // Excel は整数ポイント単位ではなく連続的に縮小する。下限は1ptとする。
            var shrunkSize = Math.Max(1.0, font.SizePt * availableWidthPt / width);
            return font with { SizePt = shrunkSize };
        }

        /// <summary>セル幅に合わせてテキストを折り返す。</summary>
        /// <remarks>
        /// 改行は LF・CRLF・CR 単独のいずれも段落区切りとして扱う。折り返し位置は書記素クラスタ
        /// (サロゲートペア・結合文字・異体字セレクタを含む、利用者が1文字と認識する単位)の境界に限る
        /// (要件4.6)。UTF-16 の1単位ごとに区切ると、「𠮷」のようなサロゲートペアの途中で改行され、
        /// 両側が文字化けする。
        /// </remarks>
        private List<string> WrapLines(FontStyle font, string text, double availableWidthPt)
        {
            var lines = new List<string>();
            var unified = text.IndexOf('\r') < 0 ? text : text.Replace("\r\n", "\n").Replace('\r', '\n');

            foreach (var paragraph in unified.Split('\n'))
            {
                if (paragraph.Length == 0)
                {
                    lines.Add(string.Empty);
                    continue;
                }

                var current = new StringBuilder();
                var elements = StringInfo.GetTextElementEnumerator(paragraph);
                while (elements.MoveNext())
                {
                    var element = elements.GetTextElement();
                    if (current.Length > 0
                        && _fontMetrics.MeasureTextWidth(font, current.ToString() + element) > availableWidthPt)
                    {
                        lines.Add(current.ToString());
                        current.Clear();
                    }

                    current.Append(element);
                }

                lines.Add(current.ToString());
            }

            return lines.Count == 0 ? new List<string> { string.Empty } : lines;
        }

        // ---------------------------------------------------------------------
        // 座標計算
        // ---------------------------------------------------------------------

        /// <summary>ページ内の各行/列の開始オフセット(論理座標、ポイント)を求める。</summary>
        private static double[] BuildOffsets(IReadOnlyList<int> indices, Func<int, double> sizeOf)
        {
            var offsets = new double[indices.Count + 1];
            for (var i = 0; i < indices.Count; i++)
            {
                offsets[i + 1] = offsets[i] + sizeOf(indices[i]);
            }

            return offsets;
        }

        private static Dictionary<int, int> BuildIndex(IReadOnlyList<int> indices)
        {
            var map = new Dictionary<int, int>(indices.Count);
            for (var i = 0; i < indices.Count; i++)
            {
                map[indices[i]] = i;
            }

            return map;
        }

        /// <summary>論理座標を、余白と拡大縮小率を適用した用紙座標へ変換する。</summary>
        private RectPt ToPageRect(double left, double top, double right, double bottom) =>
            RectPt.FromBounds(
                _originX + (left * _scale),
                _originY + (top * _scale),
                _originX + (right * _scale),
                _originY + (bottom * _scale));

        private RectPt CellRect(
            int row,
            int column,
            IReadOnlyDictionary<int, int> rowIndex,
            IReadOnlyDictionary<int, int> columnIndex,
            double[] rowOffsets,
            double[] columnOffsets)
        {
            var r = rowIndex[row];
            var c = columnIndex[column];
            return ToPageRect(columnOffsets[c], rowOffsets[r], columnOffsets[c + 1], rowOffsets[r + 1]);
        }

        /// <summary>
        /// 結合範囲のうち、このページに見えている部分の矩形を返す。
        /// ページ上に1行/1列も含まれない場合は null。
        /// </summary>
        private RectPt? TryGetMergedRect(
            CellRange range,
            IReadOnlyDictionary<int, int> rowIndex,
            IReadOnlyDictionary<int, int> columnIndex,
            double[] rowOffsets,
            double[] columnOffsets)
        {
            var (firstRowPos, lastRowPos) = FindVisibleSpan(range.FirstRow, range.LastRow, rowIndex);
            if (firstRowPos < 0)
            {
                return null;
            }

            var (firstColPos, lastColPos) = FindVisibleSpan(range.FirstColumn, range.LastColumn, columnIndex);
            if (firstColPos < 0)
            {
                return null;
            }

            return ToPageRect(
                columnOffsets[firstColPos],
                rowOffsets[firstRowPos],
                columnOffsets[lastColPos + 1],
                rowOffsets[lastRowPos + 1]);
        }

        /// <summary>
        /// 範囲 [first, last] のうち、ページ上に存在する指標の位置(連番)の最小・最大を返す。
        /// </summary>
        /// <remarks>
        /// 印刷タイトルの繰り返しにより、ページ上の指標は必ずしも連続しない。
        /// 結合範囲が非連続な位置にまたがる場合は、見えている範囲全体を1つの矩形として扱う。
        /// 走査する指標数には<see cref="ResolveColumnEdge"/>と同じ<see cref="MaxSpanCells"/>の
        /// 上限を設ける(code-reviewer指摘)。
        /// </remarks>
        private static (int First, int Last) FindVisibleSpan(
            int first, int last, IReadOnlyDictionary<int, int> index)
        {
            var minPos = int.MaxValue;
            var maxPos = -1;
            var boundedLast = Math.Min(last, first + MaxSpanCells);

            for (var value = first; value <= boundedLast; value++)
            {
                if (!index.TryGetValue(value, out var pos))
                {
                    continue;
                }

                if (pos < minPos) { minPos = pos; }
                if (pos > maxPos) { maxPos = pos; }
            }

            return maxPos < 0 ? (-1, -1) : (minPos, maxPos);
        }
    }
}
