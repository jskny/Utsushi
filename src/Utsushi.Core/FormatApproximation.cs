namespace Utsushi.Core
{
    /// <summary>
    /// 書式の近似の種類(excel2pdf-consolidation 要件7.2)。描画は欠けないが、Excelと異なる見た目で描く書式を表す。
    /// </summary>
    public enum FormatApproximationKind
    {
        /// <summary>セル・図形の文字の方向・回転・縦書き(横書きで描く)。要件7.2(a)。</summary>
        TextRotation,

        /// <summary>横位置の均等割り付け・両端揃え・繰り返し・選択範囲内で中央(別の配置で描く)。要件7.2(b)。</summary>
        HorizontalAlignment,

        /// <summary>縦位置の均等割り付け・両端揃え(中央揃えで描く)。要件7.2(b)。</summary>
        VerticalAlignment,

        /// <summary>右揃え・中央揃えのセルのインデント(左に足して描く)。要件7.2(b)。</summary>
        Indent,

        /// <summary>条件付き書式(条件によらず元の書式で描く)。要件7.2(c)。</summary>
        ConditionalFormatting,

        /// <summary>一部またはすべての文字の書式がセルの書式と異なるセル(セル全体をセルの書式で描く)。要件7.2(d)。</summary>
        RichText,

        /// <summary>解釈できない表示形式・表示形式で表せない値(標準の表示形式で表示する)。要件7.2(e)。</summary>
        NumberFormat,

        /// <summary>セルの網かけ(背景色の単色で描く)。要件7.2(f)。</summary>
        CellFillPattern,

        /// <summary>セルのグラデーション(塗りつぶしなしで描く)。要件7.2(f)。</summary>
        CellFillGradient,

        /// <summary>図形の塗りつぶしの近似・欠落。要件7.2(g)。</summary>
        ShapeFill,

        /// <summary>図形の文字のスタイル(<c>fontRef</c>)が指す書体(使わない)。要件7.2(g)。</summary>
        ShapeFont,

        /// <summary>反転した画像(反転せずに描く)。要件7.2(h)。</summary>
        ImageFlip,

        /// <summary>45〜135度・225〜315度回転した画像(位置・大きさがずれうる)。要件7.2(h)。</summary>
        ImageRotation,

        /// <summary>複数行のヘッダー/フッター(1行で描く)。要件7.2(i)。</summary>
        MultiLineHeaderFooter,

        /// <summary>改ページをまたぐ結合セル(両方のページに同じ文字を描く)。要件7.2(j)。</summary>
        MergedCellAcrossPageBreak,

        /// <summary>改ページをまたぐ画像・図形(左上が属するページにだけ描く)。要件7.2(j)。</summary>
        DrawingAcrossPageBreak,

        /// <summary>最後の行・列が非表示の結合セル(外周の罫線の一部を描かない)。要件7.2(j)。</summary>
        MergedCellHiddenEdge,
    }

    /// <summary>
    /// 書式の近似1件(excel2pdf-consolidation 要件7.1)。描画は欠けないがExcelと見た目が異なる書式と、その位置。
    /// </summary>
    /// <param name="Kind">種類。</param>
    /// <param name="SheetName">シート名。</param>
    /// <param name="Location">セル(1×1の範囲)・セル範囲・要素のアンカー。ヘッダー/フッターでは <c>null</c>。</param>
    /// <param name="Detail">書式の値(例: <c>textRotation=255</c>)。セル・図形の文字列は含まない。</param>
    /// <param name="Message">人が読む説明。セル・図形の文字列は含まない。</param>
    public sealed record FormatApproximation(
        FormatApproximationKind Kind,
        string SheetName,
        CellRange? Location,
        string? Detail,
        string Message)
    {
        /// <summary>1行の説明(種類・シート・位置・説明)。</summary>
        public override string ToString()
        {
            var location = $"シート '{DiagnosticText.Normalize(SheetName)}'";
            if (Location is { } range)
            {
                location += " " + (range.TopLeft == range.BottomRight ? range.TopLeft.ToString() : range.ToString());
            }

            var detail = Detail is null ? string.Empty : $" ({DiagnosticText.Normalize(Detail)})";
            return $"[{Kind}] {location}: {DiagnosticText.Normalize(Message, int.MaxValue)}{detail}";
        }
    }
}
