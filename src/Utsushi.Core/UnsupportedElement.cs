namespace Utsushi.Core
{
    /// <summary>
    /// 未対応要素の種類(excel2pdf-consolidation 要件2.7)。描画が欠ける(出力されない)要素を表す。
    /// </summary>
    public enum UnsupportedElementKind
    {
        /// <summary>グラフ(グラフを参照する図表枠)。要件2.7(a)。</summary>
        Chart,

        /// <summary>グラフ以外の図表枠・その他の非対応の描画オブジェクト・解決できない描画パート。要件2.7(a)。</summary>
        UnsupportedDrawingObject,

        /// <summary>絶対位置(<c>absoluteAnchor</c>)で配置された描画オブジェクト。要件2.7(a)。</summary>
        UnsupportedAnchor,

        /// <summary>非対応プリセット・自由形状(<c>custGeom</c>)の図形・接続線。要件2.7(b)。</summary>
        UnsupportedShapePreset,

        /// <summary>出力できない子要素を含むグループ(グループ全体を出力しない)。要件2.7(c)。</summary>
        UnsupportedGroupChild,

        /// <summary>非対応形式、または内容が形式と一致しない画像。要件2.7(d)。</summary>
        UnsupportedImageFormat,

        /// <summary>参照先が無い・空・画像でない画像。要件2.7(d)。</summary>
        MissingImagePart,

        /// <summary>枚数・大きさ・文字数・入れ子・結合セル数の上限の超過。どの上限かは <see cref="UnsupportedElement.Detail"/> に入る。要件2.7(e)。</summary>
        LimitExceeded,

        /// <summary>外部参照。要件2.7(f)。</summary>
        ExternalReference,

        /// <summary>コメント・フォームコントロールなどのレガシー図形(VML)。要件2.7(f)。</summary>
        LegacyDrawing,

        /// <summary>位置・大きさの情報が欠けている・壊れているために読み飛ばす描画オブジェクト。要件2.7(g)。</summary>
        MalformedDrawingObject,

        /// <summary>ヘッダー/フッターの画像(<c>legacyDrawingHF</c>)。要件2.7(h)。</summary>
        HeaderFooterImage,
    }

    /// <summary>
    /// 変換で出力されない(または一部が欠ける)ブック内の要素1件(excel2pdf-consolidation 要件2.4)。
    /// </summary>
    /// <param name="Kind">種類。</param>
    /// <param name="SheetName">シート名。外部参照のようにブック全体に関わる要素では <c>null</c>。</param>
    /// <param name="Anchor">アンカーの左上のセル。分からない場合は <c>null</c>。</param>
    /// <param name="Detail">補足の情報(プリセット名・画像の形式・上限の種類と値など)。セル・図形の文字列は含まない。</param>
    /// <param name="Message">人が読む説明。セル・図形の文字列は含まない。</param>
    public sealed record UnsupportedElement(
        UnsupportedElementKind Kind,
        string? SheetName,
        CellAddress? Anchor,
        string? Detail,
        string Message)
    {
        /// <summary>1行の説明(種類・シート・位置・説明)。</summary>
        public override string ToString()
        {
            var location = SheetName is null ? "ブック" : $"シート '{SheetName}'";
            if (Anchor is { } anchor)
            {
                location += " " + anchor.ToString();
            }

            return $"[{Kind}] {location}: {Message}";
        }
    }
}
