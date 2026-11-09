using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Layout.Model;

namespace Utsushi
{
    /// <summary>
    /// 診断の結果(<see cref="Excel2Pdf.Check"/>。excel2pdf-consolidation 要件4)。PDFを出力せずに、表示が崩れる原因を一覧で返す。
    /// </summary>
    /// <param name="FitIssues">文字がセルに収まらない箇所(中核仕様の要件13)。</param>
    /// <param name="FitIssuesTruncated"><paramref name="FitIssues"/> を件数の上限で打ち切ったかどうか。</param>
    /// <param name="UnsupportedElements">PDFに出力されない要素(要件2)。未対応要素の扱いによらず、例外にせず返す。</param>
    /// <param name="UnsupportedElementsTruncated"><paramref name="UnsupportedElements"/> を件数の上限で打ち切ったかどうか。</param>
    /// <param name="FontSubstitutions">テンプレートのフォントが無く、別の書体で描いたフォント(要件3)。代替を許可した場合だけ起こる。</param>
    public sealed record Excel2PdfDiagnostics(
        IReadOnlyList<FitIssue> FitIssues,
        bool FitIssuesTruncated,
        IReadOnlyList<UnsupportedElement> UnsupportedElements,
        bool UnsupportedElementsTruncated,
        IReadOnlyList<FontSubstitution> FontSubstitutions)
    {
        /// <summary>問題(文字の収まらない箇所・未対応要素・フォントの代替)が1つ以上あるかどうか(要件4.3)。</summary>
        public bool HasIssues => FitIssues.Count > 0 || UnsupportedElements.Count > 0 || FontSubstitutions.Count > 0;
    }

    /// <summary>フォントの代替1件(要件3.1)。</summary>
    /// <param name="RequestedFamily">テンプレートが指定したフォント名。</param>
    /// <param name="Bold">太字かどうか。</param>
    /// <param name="Italic">斜体かどうか。</param>
    /// <param name="ActualFamily">実際に使った書体のファミリ名(書体のデータが持つ名前。同梱のフォントでは英語名になりうる)。</param>
    public sealed record FontSubstitution(string RequestedFamily, bool Bold, bool Italic, string ActualFamily)
    {
        /// <inheritdoc />
        public override string ToString()
        {
            var style = (Bold, Italic) switch
            {
                (true, true) => "(太字・斜体)",
                (true, false) => "(太字)",
                (false, true) => "(斜体)",
                _ => string.Empty,
            };
            return RequestedFamily + style + " → " + ActualFamily;
        }
    }
}
