using System;
using System.Collections.Generic;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.HeaderFooter;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Layout;

/// <summary>
/// <see cref="IReportLayoutEngine"/> の既定実装。
/// </summary>
/// <remarks>
/// <para>
/// Excel のレイアウトルールのうち、自社帳票の再現に必要な範囲を実装する:
/// 印刷範囲のクリッピング(要件3.1)、手動/自動改ページ(要件3.2, 3.3)、
/// 印刷タイトルの繰り返し(要件3.4)、印刷順序(要件3.5)、
/// 結合セル・罫線・背景(要件4.2, 4.3)、テキスト配置(要件4.1, 4.4, 2.5)。
/// </para>
/// <para>
/// 出力する座標はページ左上原点・ポイント単位で、余白と拡大縮小率を適用済みである。
/// Rendering レイヤーは追加の座標変換を行わない。
/// </para>
/// </remarks>
public sealed class ReportLayoutEngine : IReportLayoutEngine
{
    private readonly IFontMetricsProvider _fontMetrics;
    private readonly Func<DateTime> _clock;

    /// <param name="fontMetrics">テキスト配置に使うフォントメトリクス。</param>
    /// <param name="clock">
    /// ヘッダー/フッターの <c>&amp;D</c>(日付)・<c>&amp;T</c>(時刻)に使う現在時刻。
    /// null の場合は <see cref="DateTime.Now"/>。テストで固定するために差し替えられるようにしている。
    /// </param>
    public ReportLayoutEngine(IFontMetricsProvider fontMetrics, Func<DateTime>? clock = null)
    {
        _fontMetrics = fontMetrics ?? throw new ArgumentNullException(nameof(fontMetrics));
        _clock = clock ?? (() => DateTime.Now);
    }

    /// <inheritdoc />
    public PagedLayout Compute(ReportModel report)
    {
        if (report is null)
        {
            throw new ArgumentNullException(nameof(report));
        }

        var definition = report.Definition;
        var sheet = report.Sheet;
        var pageSetup = sheet.PageSetup;

        var (paperWidthPt, paperHeightPt) = pageSetup.PaperSizePt;
        var printableWidthPt = paperWidthPt - pageSetup.Margins.LeftPt - pageSetup.Margins.RightPt;
        var printableHeightPt = paperHeightPt - pageSetup.Margins.TopPt - pageSetup.Margins.BottomPt;

        if (printableWidthPt <= 0 || printableHeightPt <= 0)
        {
            throw new LayoutComputationException(
                $"余白が用紙サイズを超えているため印字可能領域が確保できません"
                + $"(用紙 {paperWidthPt:0.#}x{paperHeightPt:0.#}pt、"
                + $"余白 左{pageSetup.Margins.LeftPt:0.#} 右{pageSetup.Margins.RightPt:0.#} "
                + $"上{pageSetup.Margins.TopPt:0.#} 下{pageSetup.Margins.BottomPt:0.#}pt)。",
                definition.ReportCode,
                sheet.Name);
        }

        // 複数の印刷範囲はそれぞれ独立したページ群になる(要件3.6)。
        var printRanges = ResolvePrintRanges(sheet, definition);
        var pages = new List<PageLayout>();

        foreach (var printRange in printRanges)
        {
            pages.AddRange(ComputePagesForRange(
                report, printRange, printableWidthPt, printableHeightPt, firstPageNumber: pages.Count + 1));
        }

        if (pages.Count == 0)
        {
            throw new LayoutComputationException(
                $"シート '{sheet.Name}' から出力可能なページがありませんでした。",
                definition.ReportCode,
                sheet.Name);
        }

        // 総ページ数が確定してからでないと &N(総ページ数)を展開できないため、
        // ヘッダー/フッターは全ページを組み立てたあとに付け足す(要件3.8)。
        pages = AppendHeadersAndFooters(report, pages);

        return new PagedLayout(pages, definition.ReportCode, sheet.Name);
    }

    /// <summary>1つの印刷範囲に対するページ群を計算する。</summary>
    private List<PageLayout> ComputePagesForRange(
        ReportModel report,
        CellRange printRange,
        double printableWidthPt,
        double printableHeightPt,
        int firstPageNumber)
    {
        var definition = report.Definition;
        var sheet = report.Sheet;
        var pageSetup = sheet.PageSetup;

        var grid = SheetGrid.Create(sheet, printRange, definition.MaxDigitWidthPx);

        // 印刷タイトルは各ページの先頭に繰り返されるため、本文の流し込みからは除外する(要件3.4)。
        var titleRows = ResolveTitleRows(grid, pageSetup.PrintTitles);
        var titleColumns = ResolveTitleColumns(grid, pageSetup.PrintTitles);
        var titleRowSet = new HashSet<int>(titleRows);
        var titleColumnSet = new HashSet<int>(titleColumns);

        var bodyRows = grid.Rows.Where(r => !titleRowSet.Contains(r)).ToList();
        var bodyColumns = grid.Columns.Where(c => !titleColumnSet.Contains(c)).ToList();

        var titleHeightPt = grid.SumRowHeights(titleRows);
        var titleWidthPt = grid.SumColumnWidths(titleColumns);

        var scale = ResolveScale(
            pageSetup.Scaling, grid, bodyRows, bodyColumns,
            printableWidthPt, printableHeightPt, titleWidthPt, titleHeightPt);

        // 論理(シート)座標での印字可能領域。拡大縮小率で割ることで、
        // 「縮小するほど1ページに多く入る」という Excel の挙動を再現する。
        var availableWidthPt = (printableWidthPt / scale) - titleWidthPt;
        var availableHeightPt = (printableHeightPt / scale) - titleHeightPt;

        if (availableWidthPt <= 0 || availableHeightPt <= 0)
        {
            throw new LayoutComputationException(
                "印刷タイトルの行/列だけで印字可能領域を使い切るため、本文を配置できません。",
                definition.ReportCode,
                sheet.Name);
        }

        var rowBands = PageBandCalculator.Split(
            bodyRows, grid.GetRowHeightPt, availableHeightPt, pageSetup.ManualRowBreaks);
        var columnBands = PageBandCalculator.Split(
            bodyColumns, grid.GetColumnWidthPt, availableWidthPt, pageSetup.ManualColumnBreaks);

        return BuildPages(
            report, grid, titleRows, titleColumns, rowBands, columnBands, scale, pageSetup, firstPageNumber);
    }

    /// <summary>
    /// 確定したページ群にヘッダー/フッターの描画命令を付け足す(要件3.7〜3.9)。
    /// </summary>
    private List<PageLayout> AppendHeadersAndFooters(ReportModel report, List<PageLayout> pages)
    {
        var headerFooter = report.Sheet.PageSetup.HeaderFooter;
        if (headerFooter.IsEmpty)
        {
            return pages;
        }

        var margins = report.Sheet.PageSetup.Margins;
        var timestamp = _clock();
        var result = new List<PageLayout>(pages.Count);

        foreach (var page in pages)
        {
            var builder = new HeaderFooterCommandBuilder(
                _fontMetrics, margins, page.WidthPt, page.HeightPt, page.ScaleFactor);

            var context = new HeaderFooterContext(
                page.PageNumber,
                pages.Count,
                report.Sheet.Name,
                report.Definition.ReportCode,
                timestamp,
                report.DefaultFont);

            var commands = builder.Build(headerFooter, context);
            if (commands.Count == 0)
            {
                result.Add(page);
                continue;
            }

            var merged = new List<DrawCommand>(page.Commands.Count + commands.Count);
            merged.AddRange(page.Commands);
            merged.AddRange(commands);
            result.Add(page with { Commands = merged });
        }

        return result;
    }

    /// <summary>
    /// 印刷範囲を決定する(要件3.1, 3.6)。
    /// 帳票定義の上書き &gt; Excelの印刷範囲設定 &gt; 使用範囲 の優先順とする。
    /// </summary>
    /// <remarks>
    /// Excel は1シートに複数の印刷範囲を設定でき、それぞれが独立したページ群として印刷される。
    /// 範囲を包含する1つの矩形として扱うと、範囲の間にある不要なセルまで出力されてしまうため、
    /// ここでは範囲を定義順のまま返す。
    /// </remarks>
    private static IReadOnlyList<CellRange> ResolvePrintRanges(SheetModel sheet, ReportDefinition definition)
    {
        if (definition.PrintAreaOverride is { } overridden)
        {
            return new[] { overridden };
        }

        var printAreas = sheet.PageSetup.PrintAreas;
        if (printAreas.Count > 0)
        {
            return printAreas;
        }

        var used = sheet.GetUsedRange();
        if (used is { } usedRange)
        {
            return new[] { usedRange };
        }

        throw new LayoutComputationException(
            $"シート '{sheet.Name}' に印刷対象のセルがありません。",
            definition.ReportCode,
            sheet.Name);
    }

    private static List<int> ResolveTitleRows(SheetGrid grid, PrintTitles titles)
    {
        if (!titles.HasRows)
        {
            return new List<int>();
        }

        return grid.Rows.Where(r => r >= titles.FirstRow!.Value && r <= titles.LastRow!.Value).ToList();
    }

    private static List<int> ResolveTitleColumns(SheetGrid grid, PrintTitles titles)
    {
        if (!titles.HasColumns)
        {
            return new List<int>();
        }

        return grid.Columns.Where(c => c >= titles.FirstColumn!.Value && c <= titles.LastColumn!.Value).ToList();
    }

    /// <summary>
    /// 拡大縮小率を決定する。
    /// 「次のページ数に合わせて印刷」が有効な場合はページ数に収まる倍率を算出し、
    /// そうでない場合は固定倍率を使う。
    /// </summary>
    private static double ResolveScale(
        PageScaling scaling,
        SheetGrid grid,
        IReadOnlyList<int> bodyRows,
        IReadOnlyList<int> bodyColumns,
        double printableWidthPt,
        double printableHeightPt,
        double titleWidthPt,
        double titleHeightPt)
    {
        if (!scaling.IsFitToPage)
        {
            return scaling.ScaleFactor;
        }

        var scale = 1.0;

        if (scaling.FitToWidth is { } fitWidth && fitWidth > 0)
        {
            var contentWidth = grid.SumColumnWidths(bodyColumns) + (titleWidthPt * fitWidth);
            if (contentWidth > 0)
            {
                scale = Math.Min(scale, printableWidthPt * fitWidth / contentWidth);
            }
        }

        if (scaling.FitToHeight is { } fitHeight && fitHeight > 0)
        {
            var contentHeight = grid.SumRowHeights(bodyRows) + (titleHeightPt * fitHeight);
            if (contentHeight > 0)
            {
                scale = Math.Min(scale, printableHeightPt * fitHeight / contentHeight);
            }
        }

        // Excel は「ページ数に合わせる」場合でも100%を超えて拡大しない。
        return Math.Min(1.0, scale);
    }

    /// <summary>
    /// 行帯・列帯の組み合わせを印刷順序に従って並べ、各ページの描画命令を生成する(要件3.5)。
    /// </summary>
    private List<PageLayout> BuildPages(
        ReportModel report,
        SheetGrid grid,
        IReadOnlyList<int> titleRows,
        IReadOnlyList<int> titleColumns,
        IReadOnlyList<IReadOnlyList<int>> rowBands,
        IReadOnlyList<IReadOnlyList<int>> columnBands,
        double scale,
        PageSetupModel pageSetup,
        int firstPageNumber)
    {
        var pages = new List<PageLayout>();
        var (paperWidthPt, paperHeightPt) = pageSetup.PaperSizePt;

        var order = pageSetup.PageOrder == PageOrder.OverThenDown
            ? EnumerateOverThenDown(rowBands.Count, columnBands.Count)
            : EnumerateDownThenOver(rowBands.Count, columnBands.Count);

        var pageNumber = firstPageNumber;
        foreach (var (rowBandIndex, columnBandIndex) in order)
        {
            var bodyRows = rowBands[rowBandIndex];
            var bodyColumns = columnBands[columnBandIndex];

            // 本文が空の帯しかない場合でも、タイトル行/列があれば1ページは出力する。
            if (bodyRows.Count == 0 && bodyColumns.Count == 0 && titleRows.Count == 0 && titleColumns.Count == 0)
            {
                continue;
            }

            var pageColumns = Concat(titleColumns, bodyColumns);
            var pageRows = Concat(titleRows, bodyRows);

            var commands = new PageCommandBuilder(report, grid, _fontMetrics, scale, pageSetup.Margins)
                .Build(pageRows, pageColumns);

            pages.Add(new PageLayout(
                pageSetup.Paper,
                pageSetup.Orientation,
                paperWidthPt,
                paperHeightPt,
                commands,
                pageNumber,
                RangeOf(bodyRows),
                RangeOf(bodyColumns),
                scale));

            pageNumber++;
        }

        return pages;
    }

    /// <summary>上から下、そのあと右へ(Excel既定)。</summary>
    private static IEnumerable<(int RowBand, int ColumnBand)> EnumerateDownThenOver(int rowBands, int columnBands)
    {
        for (var c = 0; c < columnBands; c++)
        {
            for (var r = 0; r < rowBands; r++)
            {
                yield return (r, c);
            }
        }
    }

    /// <summary>左から右、そのあと下へ。</summary>
    private static IEnumerable<(int RowBand, int ColumnBand)> EnumerateOverThenDown(int rowBands, int columnBands)
    {
        for (var r = 0; r < rowBands; r++)
        {
            for (var c = 0; c < columnBands; c++)
            {
                yield return (r, c);
            }
        }
    }

    private static List<int> Concat(IReadOnlyList<int> first, IReadOnlyList<int> second)
    {
        var result = new List<int>(first.Count + second.Count);
        result.AddRange(first);
        result.AddRange(second);
        return result;
    }

    private static (int First, int Last) RangeOf(IReadOnlyList<int> indices) =>
        indices.Count == 0 ? (0, 0) : (indices[0], indices[indices.Count - 1]);
}
