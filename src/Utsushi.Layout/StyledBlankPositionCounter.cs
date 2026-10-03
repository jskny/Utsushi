using Utsushi.Core.Exceptions;

namespace Utsushi.Layout
{
    /// <summary>
    /// ファイルにセルが無い位置のうち、行・列・ブックの標準の書式(要件1.10)で塗りつぶし・罫線を描いた位置の数を、
    /// 文書全体(全印刷範囲・全ページ)で数え、上限を超えたら中止する。
    /// </summary>
    /// <remarks>
    /// セルの描画命令の数は、これまでファイル内のセルの数(<c>MaxCellsPerSheet</c>)で抑えられていた。行・列の書式は
    /// <c>col</c> 要素1つ(数十バイト)で 16,384 列ぶんの位置に効くため、そのまま描くと、小さなファイルで印刷範囲の大きさ
    /// (<c>MaxPrintRangeCells</c>)に比例する描画命令・メモリ・PDFサイズを作らせることができる(security-reviewer指摘)。
    /// 印刷タイトルとしてページごとに繰り返し描く位置も、描くたびに数える。
    /// </remarks>
    internal sealed class StyledBlankPositionCounter
    {
        /// <summary>
        /// 上限。A4縦の1ページは多くても数千セルのため、数十ページの帳票の全面に行・列の書式があっても収まる。
        /// </summary>
        internal const int MaxStyledBlankPositions = 200_000;

        private readonly string _reportCode;
        private readonly string _sheetName;
        private readonly int _limit;
        private int _count;

        public StyledBlankPositionCounter(string reportCode, string sheetName, int limit = MaxStyledBlankPositions)
        {
            _reportCode = reportCode;
            _sheetName = sheetName;
            _limit = limit;
        }

        /// <summary>1か所を数える。上限を超えたら <see cref="LayoutComputationException"/>。</summary>
        public void Count()
        {
            _count++;
            if (_count > _limit)
            {
                throw new LayoutComputationException(
                    $"行全体・列全体の書式で塗りつぶし・罫線を描く、セルの無い位置が上限({_limit:N0}か所)を超えました。"
                    + "印刷範囲を狭めるか、行全体・列全体に設定した書式を見直してください。",
                    _reportCode,
                    _sheetName);
            }
        }
    }
}
