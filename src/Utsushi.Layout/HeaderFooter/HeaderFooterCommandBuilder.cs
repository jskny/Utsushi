using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout.HeaderFooter
{
    /// <summary>
    /// ページヘッダー/フッターの描画命令を組み立てる(要件3.7〜3.9)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// ヘッダーは「用紙上端からヘッダー余白のぶん下がった位置」に上端を合わせ、
    /// フッターは「用紙下端からフッター余白のぶん上がった位置」に下端を合わせる。
    /// これは Excel がヘッダー/フッターを本文の印字領域ではなく上下の余白の中に置く挙動に対応する。
    /// </para>
    /// <para>
    /// 横方向は、左セクションを左余白に左揃え、右セクションを右余白に右揃え、
    /// 中央セクションを左右余白の中央に配置する。
    /// </para>
    /// </remarks>
    internal sealed class HeaderFooterCommandBuilder
    {
        private readonly IFontMetricsProvider _fontMetrics;
        private readonly PageMargins _margins;
        private readonly double _pageWidthPt;
        private readonly double _pageHeightPt;
        private readonly double _scale;

        public HeaderFooterCommandBuilder(
            IFontMetricsProvider fontMetrics,
            PageMargins margins,
            double pageWidthPt,
            double pageHeightPt,
            double scale)
        {
            _fontMetrics = fontMetrics;
            _margins = margins;
            _pageWidthPt = pageWidthPt;
            _pageHeightPt = pageHeightPt;
            _scale = scale;
        }

        /// <summary>1ページ分のヘッダーとフッターの描画命令を返す。</summary>
        public IReadOnlyList<DrawCommand> Build(HeaderFooterModel headerFooter, HeaderFooterContext context)
        {
            var commands = new List<DrawCommand>();
            if (headerFooter.IsEmpty)
            {
                return commands;
            }

            // 「文書と一緒に拡大縮小する」が有効な場合のみ、拡大縮小率を文字サイズに適用する。
            var fontScale = headerFooter.ScaleWithDocument ? _scale : 1.0;

            AppendSections(
                commands,
                HeaderFooterParser.Parse(headerFooter.GetHeader(context.PageNumber), context),
                fontScale,
                isHeader: true);

            AppendSections(
                commands,
                HeaderFooterParser.Parse(headerFooter.GetFooter(context.PageNumber), context),
                fontScale,
                isHeader: false);

            return commands;
        }

        private void AppendSections(
            List<DrawCommand> commands, IReadOnlyList<HeaderFooterPart> parts, double fontScale, bool isHeader)
        {
            foreach (var part in parts)
            {
                AppendPart(commands, part, fontScale, isHeader);
            }
        }

        private void AppendPart(List<DrawCommand> commands, HeaderFooterPart part, double fontScale, bool isHeader)
        {
            var scaledRuns = ScaleRuns(part.Runs, fontScale);
            if (scaledRuns.Count == 0)
            {
                return;
            }

            var totalWidth = 0.0;
            var maxAscent = 0.0;
            var maxDescent = 0.0;

            foreach (var run in scaledRuns)
            {
                totalWidth += _fontMetrics.MeasureTextWidth(run.Font, run.Text);
                var metrics = _fontMetrics.GetMetrics(run.Font);
                maxAscent = Math.Max(maxAscent, metrics.AscentPt);
                maxDescent = Math.Max(maxDescent, metrics.DescentPt);
            }

            // 余白(用紙端からの物理的な距離)は「文書と一緒に拡大縮小する」の対象外。
            // 対象はヘッダー/フッターの文字サイズ(fontScaleは_fontMetricsに渡すFontStyle.SizePtにのみ適用済み)。
            var baselineY = isHeader
                ? _margins.HeaderPt + maxAscent
                : _pageHeightPt - _margins.FooterPt - maxDescent;

            var startX = ResolveStartX(part.Section, totalWidth);

            // 書式が変わるごとに別命令になるため、左から順に送り幅を足していく。
            var x = startX;
            foreach (var run in scaledRuns)
            {
                commands.Add(new TextCommand(new PointPt(x, baselineY), run.Text, run.Font, TextAnchor.Left, null));
                x += _fontMetrics.MeasureTextWidth(run.Font, run.Text);
            }
        }

        private double ResolveStartX(HeaderFooterSection section, double totalWidth) => section switch
        {
            HeaderFooterSection.Right => _pageWidthPt - _margins.RightPt - totalWidth,
            HeaderFooterSection.Center =>
                _margins.LeftPt + ((_pageWidthPt - _margins.LeftPt - _margins.RightPt - totalWidth) / 2.0),
            _ => _margins.LeftPt,
        };

        private static List<HeaderFooterRun> ScaleRuns(IReadOnlyList<HeaderFooterRun> runs, double fontScale)
        {
            var scaled = new List<HeaderFooterRun>(runs.Count);
            foreach (var run in runs)
            {
                if (run.Text.Length == 0)
                {
                    continue;
                }

                scaled.Add(Math.Abs(fontScale - 1.0) < 1e-9
                    ? run
                    : run with { Font = run.Font with { SizePt = run.Font.SizePt * fontScale } });
            }

            return scaled;
        }
    }
}
