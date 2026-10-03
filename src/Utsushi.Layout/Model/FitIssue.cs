using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Layout.Model
{
    /// <summary>文字がセルの表示領域に収まらない箇所の種類(要件13.3)。</summary>
    public enum FitIssueKind
    {
        /// <summary>はみ出し表示の文字が、値を持つ隣のセルに重なる(Excel はその手前で文字を止める)。</summary>
        OverlapsNeighborValue,

        /// <summary>はみ出し表示の文字が、ページの本文の範囲(余白の内側)で切れる。</summary>
        CutAtPageEdge,

        /// <summary>数値の表示文字列がセルの幅に収まらない(Excel では <c>####</c> と表示される)。</summary>
        NumberTooWide,

        /// <summary>切り取り表示、または結合セルの文字が、セル(結合範囲)の幅で切れる。</summary>
        Clipped,

        /// <summary>折り返した文字がセルの高さに収まらない。</summary>
        ExceedsCellHeight,
    }

    /// <summary>
    /// 文字がセルの表示領域に収まらない箇所(要件13)。PDF の描画には使わず、呼び出し元がログに残すための情報。
    /// </summary>
    /// <param name="Kind">種類。</param>
    /// <param name="Cell">セル番地(結合セルは範囲の左上)。</param>
    /// <param name="PageNumber">そのセルが最初に現れるページの番号(1始まりの通し番号。ヘッダー/フッターの先頭ページ番号は反映しない)。</param>
    /// <param name="Text">
    /// 描画する文字列(折り返し表示では改行を含む元の文字列、それ以外は改行を除いた1行)。宛名などの個人情報を含みうるため、
    /// <paramref name="Message"/> には含めていない。ログに残すかどうかは呼び出し元で決める。
    /// </param>
    /// <param name="IsSubstituted">置換キーまたはセル番地直接指定で値を置き換えたセルかどうか。</param>
    /// <param name="SubstitutionKey">置換キーで置き換えたセルなら、その置換キー。それ以外は null。</param>
    /// <param name="Message">ログに残すための説明の文。セル番地・置換キー・対処を含み、セルの文字列そのものは含まない。</param>
    public sealed record FitIssue(
        FitIssueKind Kind,
        CellAddress Cell,
        int PageNumber,
        string Text,
        bool IsSubstituted,
        string? SubstitutionKey,
        string Message)
    {
        /// <inheritdoc />
        public override string ToString() => Message;
    }

    /// <summary>文字の収まりの確認(要件13)の結果。</summary>
    /// <param name="Issues">収まらない箇所(セルが現れる順)。1件も無ければ空。</param>
    /// <param name="IsTruncated">
    /// 件数の上限(10,000件)に達し、それ以降の箇所を返していない場合は true。列幅の足りない大きな一覧表などで起きる。
    /// </param>
    public sealed record FitCheckResult(IReadOnlyList<FitIssue> Issues, bool IsTruncated);
}
