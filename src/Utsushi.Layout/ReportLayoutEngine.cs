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

namespace Utsushi.Layout
{
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
        /// <summary>
        /// 印刷範囲(複数ある場合は合計)の行数×列数の上限(要件6.9)。ページごとにセルを1つずつ走査するため、
        /// 極端に大きな印刷範囲・使用範囲(例: A1とXFD500000の2セルだけのシート)で処理が終わらなくなるのを防ぐ
        /// (security-reviewer指摘)。自社帳票(数千行×数十列)より十分大きい。
        /// </summary>
        internal const long MaxPrintRangeCells = 2_000_000;

        /// <summary>1文書あたりの出力ページ数の上限(要件6.9)。</summary>
        internal const int MaxPagesPerDocument = 5000;

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
            EnsurePrintRangeCellsWithinLimit(printRanges, MaxPrintRangeCells, definition.ReportCode, sheet.Name, pageSetup.PrintTitles);

            // 改ページ位置・印刷タイトルを付けるページ・拡大縮小率を印刷範囲ごとに先に決める。
            // 印刷範囲の外にある印刷タイトルが実際に出力されるかどうかは、ページ分割の結果で決まるため、
            // 差し込みセルの検証はこの後に行う。
            var plans = new List<RangePlan>(printRanges.Count);
            var totalPages = 0L;
            foreach (var printRange in printRanges)
            {
                var plan = PlanRange(report, printRange, printableWidthPt, printableHeightPt);
                totalPages += (long)plan.RowBands.Count * plan.ColumnBands.Count;
                EnsurePageCountWithinLimit(totalPages, MaxPagesPerDocument, definition.ReportCode, sheet.Name);
                plans.Add(plan);
            }

            ValidateRequiredFieldsAreInPrintRanges(definition, sheet, plans);
            ValidateSubstitutedCellsAreInPrintRanges(report, plans);

            var printableArea = RectPt.FromBounds(
                pageSetup.Margins.LeftPt,
                pageSetup.Margins.TopPt,
                paperWidthPt - pageSetup.Margins.RightPt,
                paperHeightPt - pageSetup.Margins.BottomPt);

            var pages = new List<PageLayout>();
            foreach (var plan in plans)
            {
                pages.AddRange(BuildPages(report, plan, pageSetup, printableArea, firstPageNumber: pages.Count + 1));
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

        /// <summary>
        /// 1つの印刷範囲の、ページ分割の計画(改ページ位置・印刷タイトルの付け方・拡大縮小率)。
        /// </summary>
        /// <param name="PrintRange">印刷範囲。</param>
        /// <param name="Grid">印刷範囲と印刷タイトルの行/列からなる格子。</param>
        /// <param name="Titles">印刷タイトル(行・列それぞれの格子上の並びと、そのうち最も後ろの番号)。</param>
        /// <param name="RowBands">本文の行帯(ページ単位)。印刷範囲の中にあるタイトル行は本文にも含む。</param>
        /// <param name="ColumnBands">本文の列帯(ページ単位)。</param>
        /// <param name="Scale">拡大縮小率。</param>
        /// <param name="MergedIndex">格子の行を登録したセル→結合範囲の索引(ページ間で使い回す)。</param>
        private sealed record RangePlan(
            CellRange PrintRange,
            SheetGrid Grid,
            TitlePlan Titles,
            IReadOnlyList<IReadOnlyList<int>> RowBands,
            IReadOnlyList<IReadOnlyList<int>> ColumnBands,
            double Scale,
            MergedCellIndex MergedIndex)
        {
            /// <summary>タイトル行を付けるページが1つでもあるかどうか。</summary>
            public bool AttachesTitleRowsAnywhere => RowBands.Any(Titles.AttachesRows);

            /// <summary>タイトル列を付けるページが1つでもあるかどうか。</summary>
            public bool AttachesTitleColumnsAnywhere => ColumnBands.Any(Titles.AttachesColumns);
        }

        /// <summary>
        /// 印刷タイトルの行/列と、それをどのページに付けるかの規則(要件3.4)。
        /// </summary>
        /// <remarks>
        /// Excel は、本文の帯の先頭がタイトルの最終行(列)より後ろにあるページにだけタイトルを付ける。
        /// タイトルが印刷範囲より上(左)にあれば全ページに付き、印刷範囲の中にあれば、タイトルそのものを
        /// 本文として印刷したページより後ろのページにだけ付く(タイトルの行を本文から取り除いたり、
        /// 1ページ目で本文より前へ移したりはしない)。
        /// </remarks>
        private sealed record TitlePlan(
            IReadOnlyList<int> Rows, int? LastRow, IReadOnlyList<int> Columns, int? LastColumn)
        {
            /// <summary>この行帯のページにタイトル行を付けるかどうか。</summary>
            public bool AttachesRows(IReadOnlyList<int> band) =>
                Rows.Count > 0 && (band.Count == 0 || band[0] > LastRow!.Value);

            /// <summary>この列帯のページにタイトル列を付けるかどうか。</summary>
            public bool AttachesColumns(IReadOnlyList<int> band) =>
                Columns.Count > 0 && (band.Count == 0 || band[0] > LastColumn!.Value);

            /// <summary>先頭が <paramref name="first"/> の行帯にタイトル行を付けるかどうか。</summary>
            public bool AttachesRowsBefore(int first) => Rows.Count > 0 && first > LastRow!.Value;

            /// <summary>先頭が <paramref name="first"/> の列帯にタイトル列を付けるかどうか。</summary>
            public bool AttachesColumnsBefore(int first) => Columns.Count > 0 && first > LastColumn!.Value;
        }

        /// <summary>1つの印刷範囲のページ分割を計画する。</summary>
        private static RangePlan PlanRange(
            ReportModel report,
            CellRange printRange,
            double printableWidthPt,
            double printableHeightPt)
        {
            var definition = report.Definition;
            var sheet = report.Sheet;
            var pageSetup = sheet.PageSetup;
            var printTitles = pageSetup.PrintTitles;

            var grid = SheetGrid.Create(sheet, printRange, definition.MaxDigitWidthPx, printTitles);

            // 印刷範囲の中にあるタイトル行/列は本文としても印刷する(要件3.4)。本文は印刷範囲の行/列だけ。
            var titles = new TitlePlan(
                ResolveTitleRows(grid, printTitles),
                printTitles.HasRows ? printTitles.LastRow : null,
                ResolveTitleColumns(grid, printTitles),
                printTitles.HasColumns ? printTitles.LastColumn : null);

            var bodyRows = grid.Rows.Where(r => r >= printRange.FirstRow && r <= printRange.LastRow).ToList();
            var bodyColumns = grid.Columns.Where(c => c >= printRange.FirstColumn && c <= printRange.LastColumn).ToList();

            var titleHeightPt = grid.SumRowHeights(titles.Rows);
            var titleWidthPt = grid.SumColumnWidths(titles.Columns);

            // 「次のページ数に合わせて印刷」では手動改ページを無視する(Excel の仕様)。
            var scaling = pageSetup.Scaling;
            var rowBreaks = scaling.IsFitToPage ? Array.Empty<int>() : (IReadOnlyCollection<int>)pageSetup.ManualRowBreaks;
            var columnBreaks = scaling.IsFitToPage ? Array.Empty<int>() : (IReadOnlyCollection<int>)pageSetup.ManualColumnBreaks;

            IReadOnlyList<IReadOnlyList<int>> SplitRows(double scale) =>
                PageBandCalculator.Split(
                    bodyRows,
                    grid.GetRowHeightPt,
                    first => (printableHeightPt / scale) - (titles.AttachesRowsBefore(first) ? titleHeightPt : 0.0),
                    rowBreaks);

            IReadOnlyList<IReadOnlyList<int>> SplitColumns(double scale) =>
                PageBandCalculator.Split(
                    bodyColumns,
                    grid.GetColumnWidthPt,
                    first => (printableWidthPt / scale) - (titles.AttachesColumnsBefore(first) ? titleWidthPt : 0.0),
                    columnBreaks);

            var resolvedScale = ResolveScale(
                scaling, grid, bodyRows, bodyColumns, printableWidthPt, printableHeightPt, SplitRows, SplitColumns);

            // タイトルを付けるページがあるのに、タイトルだけで印字可能領域を使い切る場合は本文を置けない。
            var rowTitlesNeeded = titles.Rows.Count > 0
                && (bodyRows.Count == 0 || titles.AttachesRowsBefore(bodyRows[bodyRows.Count - 1]));
            var columnTitlesNeeded = titles.Columns.Count > 0
                && (bodyColumns.Count == 0 || titles.AttachesColumnsBefore(bodyColumns[bodyColumns.Count - 1]));
            if ((rowTitlesNeeded && (printableHeightPt / resolvedScale) - titleHeightPt <= 0)
                || (columnTitlesNeeded && (printableWidthPt / resolvedScale) - titleWidthPt <= 0))
            {
                throw new LayoutComputationException(
                    "印刷タイトルの行/列だけで印字可能領域を使い切るため、本文を配置できません。",
                    definition.ReportCode,
                    sheet.Name);
            }

            var rowBands = SplitRows(resolvedScale);
            var columnBands = SplitColumns(resolvedScale);

            return new RangePlan(
                printRange,
                grid,
                titles,
                rowBands,
                columnBands,
                resolvedScale,
                MergedCellIndex.Create(sheet.MergedRanges, grid.Rows));
        }

        /// <summary>
        /// 印刷範囲の行数×列数の合計が上限以下か確認する(要件6.9。テスト用に上限を引数に取る)。
        /// 印刷タイトルの行・列はすべての印刷範囲・ページで繰り返し走査されるため、各印刷範囲の行数・列数に
        /// タイトルの行数・列数を足して数える(タイトルだけで上限を迂回されないように。security-reviewer指摘)。
        /// </summary>
        internal static void EnsurePrintRangeCellsWithinLimit(
            IReadOnlyList<CellRange> printRanges, long maxCells, string? reportCode, string sheetName, PrintTitles? titles = null)
        {
            var titleRows = titles is { HasRows: true } ? Math.Max(0L, (long)titles.LastRow!.Value - titles.FirstRow!.Value + 1) : 0L;
            var titleColumns = titles is { HasColumns: true }
                ? Math.Max(0L, (long)titles.LastColumn!.Value - titles.FirstColumn!.Value + 1)
                : 0L;

            var total = 0L;
            foreach (var range in printRanges)
            {
                total += (range.RowCount + titleRows) * (range.ColumnCount + titleColumns);
                if (total > maxCells)
                {
                    throw new LayoutComputationException(
                        $"シート '{sheetName}' の印刷範囲が大きすぎます(行数×列数の合計が上限 {maxCells:N0} を超えています)。"
                        + "印刷範囲を設定し直してください。",
                        reportCode,
                        sheetName);
                }
            }
        }

        /// <summary>出力ページ数が上限以下か確認する(要件6.9。テスト用に上限を引数に取る)。</summary>
        internal static void EnsurePageCountWithinLimit(long pageCount, int maxPages, string? reportCode, string sheetName)
        {
            if (pageCount > maxPages)
            {
                throw new LayoutComputationException(
                    $"シート '{sheetName}' の出力ページ数が上限({maxPages}ページ)を超えます。",
                    reportCode,
                    sheetName);
            }
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

            // 「先頭ページ番号」が指定されていれば、&P はその番号から数え、&N は最終ページの番号にする。
            var firstNumber = report.Sheet.PageSetup.FirstPageNumber ?? 1;

            foreach (var page in pages)
            {
                var builder = new HeaderFooterCommandBuilder(
                    _fontMetrics, margins, page.WidthPt, page.HeightPt, page.ScaleFactor);

                var context = new HeaderFooterContext(
                    firstNumber + page.PageNumber - 1,
                    firstNumber + pages.Count - 1,
                    report.Sheet.Name,
                    report.Definition.ReportCode,
                    timestamp,
                    report.DefaultFont,
                    IsFirstPage: page.PageNumber == 1);

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

            // 印刷範囲が無ければ、セルと描画オブジェクトが置かれた範囲を印刷する(要件3.10)。
            var used = UsedRangeResolver.Resolve(sheet, definition.MaxDigitWidthPx);
            if (used is { } usedRange)
            {
                return new[] { usedRange };
            }

            throw new LayoutComputationException(
                $"シート '{sheet.Name}' に印刷対象のセルがありません。",
                definition.ReportCode,
                sheet.Name);
        }

        /// <summary>
        /// 必須の置換フィールドが、実際に出力される印刷範囲(印刷タイトルを含む)の内側にあることを
        /// 確認する(要件2.4, 3.1, 3.4)。
        /// </summary>
        /// <remarks>
        /// <para>
        /// <see cref="Substitution.ICellSubstitutor"/> は印刷範囲を考慮せず値をセルにセットするため、
        /// 必須フィールドのセルが印刷範囲の外にあると、値は正しくセットされたのにPDFには出力されない
        /// という静かなデータ欠落が起こり得る。ReportDefinitionレイヤーの検証(使用範囲内かどうか)では
        /// 印刷範囲(Excel側の設定に従う場合はここで初めて解決される)まではわからないため、
        /// 印刷範囲が確定するこのタイミングで検証する。
        /// </para>
        /// <para>
        /// 印刷タイトルの行/列は印刷範囲の外でも各ページに出力される(<see cref="SheetGrid.Create"/>参照)ため、
        /// 印刷範囲だけでなく、実際にタイトルを付けるページがある印刷タイトルの行/列も「実際に出力される」対象に含めて判定する。
        /// </para>
        /// </remarks>
        private static void ValidateRequiredFieldsAreInPrintRanges(
            ReportDefinition definition, SheetModel sheet, IReadOnlyList<RangePlan> plans)
        {
            foreach (var field in definition.SubstitutionFields)
            {
                if (!field.Required)
                {
                    continue;
                }

                if (plans.Any(plan => IsRenderedByPlan(field.Cell, plan, sheet.PageSetup.PrintTitles)))
                {
                    continue;
                }

                throw new LayoutComputationException(
                    $"必須の置換キー '{field.Key}' の対象セル {field.Cell} が印刷範囲の外にあるため、"
                    + "値を設定してもPDFに出力されません。帳票定義の印刷範囲またはセル番地を見直してください。",
                    definition.ReportCode,
                    sheet.Name,
                    field.Cell);
            }
        }

        /// <summary>
        /// 空でない置換値が差し込まれたセル(任意の置換キー・セル番地直接指定を含む)が、
        /// 実際に出力される印刷範囲(印刷タイトルを含む)の内側にあることを確認する(要件2.13)。
        /// </summary>
        /// <remarks>
        /// 必須フィールドは値の有無によらず <see cref="ValidateRequiredFieldsAreInPrintRanges"/> で
        /// 先に検証している。ここでは、値を渡したのに出力されない残りの経路を塞ぐ。
        /// </remarks>
        private static void ValidateSubstitutedCellsAreInPrintRanges(ReportModel report, IReadOnlyList<RangePlan> plans)
        {
            var titles = report.Sheet.PageSetup.PrintTitles;
            foreach (var cell in report.SubstitutedCells.OrderBy(c => c.Row).ThenBy(c => c.Column))
            {
                if (IsHidden(report.Sheet, cell, report.Definition.MaxDigitWidthPx))
                {
                    throw new LayoutComputationException(
                        $"値を差し込んだセル {cell} が非表示の行または列(高さ・幅が0の行・列を含む)にあるため、"
                        + "PDFに出力されません。テンプレートの非表示設定・行の高さ・列幅またはセル番地を見直してください。",
                        report.Definition.ReportCode,
                        report.Sheet.Name,
                        cell);
                }

                if (plans.Any(plan => IsRenderedByPlan(cell, plan, titles)))
                {
                    continue;
                }

                throw new LayoutComputationException(
                    $"値を差し込んだセル {cell} が印刷範囲の外にあるため、PDFに出力されません。"
                    + "帳票定義の印刷範囲またはセル番地を見直してください。",
                    report.Definition.ReportCode,
                    report.Sheet.Name,
                    cell);
            }
        }

        /// <summary>
        /// セル(結合範囲のアンカーなら範囲全体)のすべての行、またはすべての列が印刷されない
        /// (非表示、または高さ・幅が0。<see cref="SheetGrid.PrintedRowHeightPt"/>と同じ基準)かどうか。
        /// </summary>
        private static bool IsHidden(SheetModel sheet, CellAddress cell, double maxDigitWidthPx)
        {
            var range = sheet.FindMergedRange(cell)?.Range ?? new CellRange(cell, cell);
            var lastRow = Math.Min(range.LastRow, range.FirstRow + 4096);
            var lastColumn = Math.Min(range.LastColumn, range.FirstColumn + 4096);

            var allRowsHidden = Enumerable.Range(range.FirstRow, lastRow - range.FirstRow + 1)
                .All(r => SheetGrid.IsRowNotPrinted(sheet, r));
            var allColumnsHidden = Enumerable.Range(range.FirstColumn, lastColumn - range.FirstColumn + 1)
                .All(c => SheetGrid.IsColumnNotPrinted(sheet, c, maxDigitWidthPx));
            return allRowsHidden || allColumnsHidden;
        }

        /// <summary>
        /// <paramref name="cell"/> が、印刷範囲と(実際にいずれかのページに付く)印刷タイトルを合わせた
        /// 出力対象(行・列それぞれ独立に判定)に含まれるかどうかを返す(<see cref="SheetGrid.Create"/>と対応)。
        /// </summary>
        private static bool IsRenderedByPlan(CellAddress cell, RangePlan plan, PrintTitles titles)
        {
            var printRange = plan.PrintRange;
            var rowIncluded = (cell.Row >= printRange.FirstRow && cell.Row <= printRange.LastRow)
                || (titles.HasRows && cell.Row >= titles.FirstRow!.Value && cell.Row <= titles.LastRow!.Value
                    && plan.AttachesTitleRowsAnywhere);

            var columnIncluded = (cell.Column >= printRange.FirstColumn && cell.Column <= printRange.LastColumn)
                || (titles.HasColumns && cell.Column >= titles.FirstColumn!.Value && cell.Column <= titles.LastColumn!.Value
                    && plan.AttachesTitleColumnsAnywhere);

            return rowIncluded && columnIncluded;
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

        /// <summary>「ページ数に合わせる」で下げられる倍率の下限(%)。Excel の拡大縮小率の下限と同じ。</summary>
        internal const int MinFitScalePercent = 10;

        /// <summary>
        /// 拡大縮小率を決定する。
        /// 「次のページ数に合わせて印刷」が有効な場合はページ数に収まる倍率を算出し、
        /// そうでない場合は固定倍率を使う。
        /// </summary>
        /// <remarks>
        /// ページ数に合わせる場合、Excel と同じく倍率は整数%に切り捨て、100%を超えて拡大しない。
        /// 行/列の区切り位置によっては、合計の大きさから求めた倍率でも指定のページ数に収まらないため、
        /// 実際に帯へ分割してページ数を数え、指定を超える間は1%ずつ下げる(下限は <see cref="MinFitScalePercent"/>)。
        /// </remarks>
        private static double ResolveScale(
            PageScaling scaling,
            SheetGrid grid,
            IReadOnlyList<int> bodyRows,
            IReadOnlyList<int> bodyColumns,
            double printableWidthPt,
            double printableHeightPt,
            Func<double, IReadOnlyList<IReadOnlyList<int>>> splitRows,
            Func<double, IReadOnlyList<IReadOnlyList<int>>> splitColumns)
        {
            if (!scaling.IsFitToPage)
            {
                return scaling.ScaleFactor;
            }

            var fitWidth = scaling.FitToWidth is { } w && w > 0 ? w : (int?)null;
            var fitHeight = scaling.FitToHeight is { } h && h > 0 ? h : (int?)null;

            // 本文の合計の大きさから求まる倍率は、指定のページ数に収まる倍率の上限である
            // (タイトルや区切り位置の分だけ、実際にはそれ以下になる)。ここから下げていく。
            var percent = 100;
            if (fitWidth is { } fw)
            {
                percent = Math.Min(percent, UpperBoundPercent(grid.SumColumnWidths(bodyColumns), printableWidthPt * fw));
            }

            if (fitHeight is { } fh)
            {
                percent = Math.Min(percent, UpperBoundPercent(grid.SumRowHeights(bodyRows), printableHeightPt * fh));
            }

            percent = Math.Max(MinFitScalePercent, percent);

            while (percent > MinFitScalePercent)
            {
                var scale = percent / 100.0;
                var fits = (fitWidth is not { } width || splitColumns(scale).Count <= width)
                    && (fitHeight is not { } height || splitRows(scale).Count <= height);
                if (fits)
                {
                    break;
                }

                percent--;
            }

            return percent / 100.0;
        }

        /// <summary>
        /// 合計 <paramref name="contentPt"/> を <paramref name="capacityPt"/> に収める倍率(%)の上限。整数%に切り捨てる。
        /// </summary>
        private static int UpperBoundPercent(double contentPt, double capacityPt)
        {
            if (contentPt <= 0)
            {
                return 100;
            }

            var percent = (capacityPt / contentPt * 100.0) + 1e-6;
            return percent >= 100.0 ? 100 : (int)Math.Floor(percent);
        }

        /// <summary>
        /// 行帯・列帯の組み合わせを印刷順序に従って並べ、各ページの描画命令を生成する(要件3.5)。
        /// </summary>
        private List<PageLayout> BuildPages(
            ReportModel report,
            RangePlan plan,
            PageSetupModel pageSetup,
            RectPt printableArea,
            int firstPageNumber)
        {
            var pages = new List<PageLayout>();
            var (paperWidthPt, paperHeightPt) = pageSetup.PaperSizePt;
            var grid = plan.Grid;
            var scale = plan.Scale;
            var rowBands = plan.RowBands;
            var columnBands = plan.ColumnBands;

            var order = pageSetup.PageOrder == PageOrder.OverThenDown
                ? EnumerateOverThenDown(rowBands.Count, columnBands.Count)
                : EnumerateDownThenOver(rowBands.Count, columnBands.Count);

            var pageNumber = firstPageNumber;
            foreach (var (rowBandIndex, columnBandIndex) in order)
            {
                var bodyRows = rowBands[rowBandIndex];
                var bodyColumns = columnBands[columnBandIndex];

                // 本文の帯の先頭がタイトルより後ろのページにだけ、タイトルを付ける(TitlePlan参照)。
                var titleRows = plan.Titles.AttachesRows(bodyRows) ? plan.Titles.Rows : Array.Empty<int>();
                var titleColumns = plan.Titles.AttachesColumns(bodyColumns) ? plan.Titles.Columns : Array.Empty<int>();

                // 本文が空の帯しかない場合でも、タイトル行/列があれば1ページは出力する。
                if (bodyRows.Count == 0 && bodyColumns.Count == 0 && titleRows.Count == 0 && titleColumns.Count == 0)
                {
                    continue;
                }

                var pageColumns = Concat(titleColumns, bodyColumns);
                var pageRows = Concat(titleRows, bodyRows);

                // 「ページ中央」: 本文(タイトルを含む)を印字可能領域の中央へ平行移動する。
                // 本文が印字可能領域より大きい場合は移動しない(左上をそろえたまま)。
                var offsetX = pageSetup.HorizontalCentered
                    ? Math.Max(0.0, (printableArea.Width - (grid.SumColumnWidths(pageColumns) * scale)) / 2.0)
                    : 0.0;
                var offsetY = pageSetup.VerticalCentered
                    ? Math.Max(0.0, (printableArea.Height - (grid.SumRowHeights(pageRows) * scale)) / 2.0)
                    : 0.0;

                var commands = new PageCommandBuilder(
                        report, grid, _fontMetrics, scale, pageSetup.Margins, new PointPt(offsetX, offsetY), printableArea,
                        plan.MergedIndex)
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
}
