using System;
using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
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
        private readonly PageMargins _margins;

        /// <summary>
        /// 2セルアンカー画像の幅/高さ計算で合算する列/行数の上限。対角セルにセル番地の上限
        /// (最大1,048,576行×16,384列)近くを指定する不正な入力による計算量の増大を防ぐ
        /// (security-reviewer指摘)。自社ロゴ用途でこの上限に達することは想定していない。
        /// </summary>
        private const int MaxSpanCells = 4096;

        /// <summary>画像・図形1つの表示サイズ(pt)の上限。異常に大きいEMU値に対する安全弁。</summary>
        private const double MaxDrawingObjectDimensionPt = 5000.0;

        /// <summary>図形内テキストの矩形内側の余白(pt)。セルの<see cref="ExcelUnitConverter.CellPaddingPoints"/>とは別に、図形用の小さめの値を使う。</summary>
        private const double ShapeTextPaddingPt = 4.0;

        private readonly List<DrawCommand> _fills = new();
        private readonly List<DrawCommand> _borders = new();
        private readonly List<DrawCommand> _texts = new();
        private readonly List<DrawCommand> _drawingObjects = new();
        private readonly HashSet<string> _emittedBorders = new(StringComparer.Ordinal);

        public PageCommandBuilder(
            ReportModel report,
            SheetGrid grid,
            IFontMetricsProvider fontMetrics,
            double scale,
            PageMargins margins)
        {
            _report = report;
            _sheet = report.Sheet;
            _grid = grid;
            _fontMetrics = fontMetrics;
            _scale = scale;
            _margins = margins;
        }

        /// <summary>ページに含まれる行・列からの描画命令を生成する。</summary>
        public IReadOnlyList<DrawCommand> Build(IReadOnlyList<int> rows, IReadOnlyList<int> columns)
        {
            var columnOffsets = BuildOffsets(columns, _grid.GetColumnWidthPt);
            var rowOffsets = BuildOffsets(rows, _grid.GetRowHeightPt);

            var columnIndex = BuildIndex(columns);
            var rowIndex = BuildIndex(rows);

            // 同一ページ上で同じ結合範囲を二重に描かないための記録。
            var emittedMergedRanges = new HashSet<CellRange>();

            foreach (var row in rows)
            {
                foreach (var column in columns)
                {
                    var address = new CellAddress(row, column);
                    var merged = _sheet.FindMergedRange(address);

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
                        EmitCell(merged.Anchor, anchorCell, mergedRect.Value, borders);
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

        // ---------------------------------------------------------------------
        // 画像・図形(要件9, 10)
        // ---------------------------------------------------------------------

        /// <summary>
        /// アンカー左上セルがこのページに含まれる画像・図形を、ページ座標に変換して描画命令にする。
        /// <see cref="SheetModel.DrawingObjects"/> の出現順(drawing.xmlの重なり順。要件10.3)を
        /// そのまま維持するため、画像・図形をまとめて1回だけ列挙する。改ページ位置をまたぐ
        /// 画像・図形は、アンカー左上セルが属するページにのみ全体を配置する(design.md「未決事項」の割り切り)。
        /// </summary>
        private void EmitDrawingObjects(
            IReadOnlyDictionary<int, int> rowIndex,
            IReadOnlyDictionary<int, int> columnIndex,
            double[] rowOffsets,
            double[] columnOffsets)
        {
            foreach (var drawingObject in _sheet.DrawingObjects)
            {
                if (!TryComputeDrawingObjectRect(drawingObject, rowIndex, columnIndex, rowOffsets, columnOffsets, out var rect))
                {
                    continue;
                }

                switch (drawingObject)
                {
                    case ImageModel image:
                        _drawingObjects.Add(new ImageCommand(rect, image.Data, image.ContentType));
                        break;
                    case ShapeModel shape:
                        _drawingObjects.Add(BuildShapeCommand(shape, rect));
                        break;
                }
            }
        }

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

            // 異常に大きいEMU値(またはその合算)による過大な矩形を防ぐ(security-reviewer指摘)。
            widthPt = Math.Min(widthPt, MaxDrawingObjectDimensionPt);
            heightPt = Math.Min(heightPt, MaxDrawingObjectDimensionPt);

            if (widthPt <= 0 || heightPt <= 0)
            {
                return false;
            }

            rect = ToPageRect(left, top, left + widthPt, top + heightPt);
            return !rect.IsEmpty;
        }

        /// <summary>図形の描画命令を組み立てる(要件10)。テキストの折り返し・配置はここで確定させ、回転前のローカル座標で保持する(回転はRenderingレイヤーが適用する)。</summary>
        private ShapeCommand BuildShapeCommand(ShapeModel shape, RectPt rect)
        {
            var textLines = shape.Text is { } text
                ? BuildShapeTextLines(text, rect)
                : Array.Empty<ShapeTextLine>();

            return new ShapeCommand(rect, shape.Preset, shape.AdjustmentValues, shape.RotationDegrees, shape.Fill, shape.Outline, textLines);
        }

        /// <summary>図形内テキストを矩形幅で折り返し、水平/垂直配置に基づく各行のローカル座標を確定させる(要件10.4)。</summary>
        private IReadOnlyList<ShapeTextLine> BuildShapeTextLines(ShapeTextBody text, RectPt rect)
        {
            var contentRect = RectPt.FromBounds(
                rect.Left + ShapeTextPaddingPt,
                rect.Top + ShapeTextPaddingPt,
                rect.Right - ShapeTextPaddingPt,
                rect.Bottom - ShapeTextPaddingPt);

            if (contentRect.Width <= 0 || contentRect.Height <= 0)
            {
                return Array.Empty<ShapeTextLine>();
            }

            var wrapped = WrapShapeText(text, contentRect.Width);
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
        private List<(string Text, HorizontalAlignment HAlign, FontStyle Font)> WrapShapeText(
            ShapeTextBody text, double availableWidthPt)
        {
            var lines = new List<(string, HorizontalAlignment, FontStyle)>();
            foreach (var paragraph in text.Paragraphs)
            {
                if (paragraph.Runs.Count == 0)
                {
                    lines.Add((string.Empty, paragraph.HAlign, FontStyle.Default));
                    continue;
                }

                var font = paragraph.Runs[0].Font;
                var paragraphText = string.Concat(paragraph.Runs.Select(run => run.Text));
                foreach (var line in WrapLines(font, paragraphText, availableWidthPt))
                {
                    lines.Add((line, paragraph.HAlign, font));
                }
            }

            return lines;
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
            _sheet.IsColumnHidden(column)
                ? 0.0
                : ExcelUnitConverter.ColumnWidthToPoints(_sheet.GetColumnWidth(column), _report.Definition.MaxDigitWidthPx);

        /// <summary>印刷範囲によらない、シート上の実際の行高(pt)。非表示行は0。</summary>
        private double RawRowHeightPt(int row) => _sheet.IsRowHidden(row) ? 0.0 : _sheet.GetRowHeight(row);

        private void EmitCell(CellAddress address, CellModel? cell, RectPt rect, BorderSet? bordersOverride = null)
        {
            if (rect.IsEmpty)
            {
                return;
            }

            var style = cell?.Style ?? CellStyle.Default;

            if (!style.BackgroundColor.IsTransparent)
            {
                _fills.Add(new FillRectCommand(rect, style.BackgroundColor));
            }

            EmitBorders(rect, bordersOverride ?? style.Borders);

            var text = cell?.DisplayValue;
            if (!string.IsNullOrEmpty(text))
            {
                EmitText(address, cell!, style, rect, text!);
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

        private BorderEdge ResolveColumnEdge(int column, int firstRow, int lastRow, Func<BorderSet, BorderEdge> selector)
        {
            for (var row = firstRow; row <= lastRow; row++)
            {
                var edge = selector(_sheet.GetCell(new CellAddress(row, column))?.Style.Borders ?? BorderSet.None);
                if (edge.IsVisible)
                {
                    return edge;
                }
            }

            return BorderEdge.None;
        }

        private BorderEdge ResolveRowEdge(int row, int firstColumn, int lastColumn, Func<BorderSet, BorderEdge> selector)
        {
            for (var column = firstColumn; column <= lastColumn; column++)
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
                var offset = BorderMetrics.DoubleLineGapPt * _scale / 2.0;
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

        private void EmitText(CellAddress address, CellModel cell, CellStyle style, RectPt rect, string text)
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

            var lines = overflow == OverflowBehavior.Wrap
                ? WrapLines(scaledFont, text, contentRect.Width)
                : new List<string> { text };

            if (overflow == OverflowBehavior.Shrink && lines.Count == 1)
            {
                scaledFont = ShrinkToFit(scaledFont, lines[0], contentRect.Width);
            }

            var metrics = _fontMetrics.GetMetrics(scaledFont);
            var totalHeight = metrics.LineSpacingPt * lines.Count;
            var firstBaselineY = ResolveFirstBaselineY(style.VAlign, rect, metrics, totalHeight);

            RectPt? clipRect = overflow is OverflowBehavior.Clip or OverflowBehavior.Wrap or OverflowBehavior.Shrink
                ? rect
                : null;

            for (var i = 0; i < lines.Count; i++)
            {
                var baselineY = firstBaselineY + (metrics.LineSpacingPt * i);
                var (x, anchor) = ResolveTextOrigin(hAlign, contentRect);
                _texts.Add(new TextCommand(new PointPt(x, baselineY), lines[i], scaledFont, anchor, clipRect));
            }
        }

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
        private List<string> WrapLines(FontStyle font, string text, double availableWidthPt)
        {
            var lines = new List<string>();

            foreach (var paragraph in text.Split('\n'))
            {
                var normalized = paragraph.TrimEnd('\r');
                if (normalized.Length == 0)
                {
                    lines.Add(string.Empty);
                    continue;
                }

                var current = string.Empty;
                foreach (var c in normalized)
                {
                    var candidate = current + c;
                    if (current.Length > 0 && _fontMetrics.MeasureTextWidth(font, candidate) > availableWidthPt)
                    {
                        lines.Add(current);
                        current = c.ToString();
                        continue;
                    }

                    current = candidate;
                }

                lines.Add(current);
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
                _margins.LeftPt + (left * _scale),
                _margins.TopPt + (top * _scale),
                _margins.LeftPt + (right * _scale),
                _margins.TopPt + (bottom * _scale));

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
        /// </remarks>
        private static (int First, int Last) FindVisibleSpan(
            int first, int last, IReadOnlyDictionary<int, int> index)
        {
            var minPos = int.MaxValue;
            var maxPos = -1;

            for (var value = first; value <= last; value++)
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
