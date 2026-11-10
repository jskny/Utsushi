namespace Utsushi.Layout
{
    /// <summary>
    /// 文字がセル幅に収まらない場合の描き方。セルの書式(折り返し・縮小)と値の種類から決める(要件2.5。excel2pdf-consolidation 4.3)。
    /// </summary>
    internal enum OverflowBehavior
    {
        /// <summary>セル幅に収まるようフォントを縮小する(Excel の「縮小して全体を表示する」)。</summary>
        Shrink = 0,

        /// <summary>セル幅で切り詰める(折り返し表示のセルの数値)。</summary>
        Clip,

        /// <summary>セル内で折り返す(Excel の「折り返して全体を表示する」)。</summary>
        Wrap,

        /// <summary>はみ出しを許容し、隣接セルへ描画する(Excel の既定)。</summary>
        Overflow,
    }
}
