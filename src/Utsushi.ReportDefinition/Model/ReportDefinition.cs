using System;

namespace Utsushi.ReportDefinitions.Model
{
    /// <summary>
    /// 1回の変換の設定(文書名・シート名・最大数字幅)。帳票ごとの差異はテンプレート(Excelの書式・印刷設定)と
    /// 呼び出し元が書き込む値で表し、共通レイヤーのコードに帳票固有の分岐を持ち込まない(excel2pdf-consolidation 要件1.5)。
    /// </summary>
    /// <param name="ReportCode">文書名。PDFのタイトル・ヘッダー/フッターの <c>&amp;F</c>・エラー情報に使う(要件12.5)。</param>
    /// <param name="SheetName">変換するシートの名前。</param>
    /// <param name="MaxDigitWidthPx">
    /// 列幅(文字数単位)→ピクセル換算に用いる標準フォントの最大数字幅(96dpiのピクセル)。
    /// Excel の換算式が標準フォント依存であるため、呼び出し元が指定できるようにしている。
    /// </param>
    public sealed record ReportDefinition(
        string ReportCode,
        string SheetName,
        double MaxDigitWidthPx)
    {
        /// <summary>Excel の標準フォント(11pt)における最大数字幅の既定値(ピクセル)。</summary>
        public const double DefaultMaxDigitWidthPx = 7.0;

        /// <summary>変換の設定を作る。</summary>
        /// <param name="documentName">文書名(PDFタイトル・ヘッダー/フッターの <c>&amp;F</c>・エラー情報に使われる。要件12.5)。</param>
        /// <param name="sheetName">変換対象のシート名。</param>
        /// <param name="maxDigitWidthPx">
        /// 列幅の換算に使う最大数字幅(ピクセル)。通常は <see cref="EstimateMaxDigitWidthPx"/> でブックの標準フォントから求める。
        /// </param>
        public static ReportDefinition CreateWithoutDefinition(
            string documentName, string sheetName, double maxDigitWidthPx = DefaultMaxDigitWidthPx) =>
            new(
                documentName ?? throw new ArgumentNullException(nameof(documentName)),
                sheetName ?? throw new ArgumentNullException(nameof(sheetName)),
                maxDigitWidthPx > 0 && !double.IsNaN(maxDigitWidthPx) && !double.IsInfinity(maxDigitWidthPx)
                    ? maxDigitWidthPx
                    : throw new ArgumentOutOfRangeException(nameof(maxDigitWidthPx), maxDigitWidthPx, "最大数字幅は正の数である必要があります。"));

        /// <summary>
        /// ブックの標準フォント(名前・サイズ)から、Excelが列幅の換算に使う最大数字幅(96dpiのピクセル)を見積もる(要件12.3)。
        /// </summary>
        /// <remarks>
        /// Excelの既定の標準フォントについて、既定の列幅の表示(英語版 Calibri 11 の「8.43(64ピクセル)」、日本語版
        /// ＭＳ Ｐゴシック/游ゴシック 11 の「8.38(72ピクセル)」)と列幅の換算式から逆算した値を持つ。表に無いフォント・サイズは
        /// <see cref="DefaultMaxDigitWidthPx"/> を返す。
        /// </remarks>
        public static double EstimateMaxDigitWidthPx(string? fontName, double sizePt)
        {
            // NaN のサイズも表の対象外として既定値にする(比較が常に false になるため、否定形で判定する)。
            if (!(Math.Abs(sizePt - 11.0) <= 0.01) || fontName is null)
            {
                return DefaultMaxDigitWidthPx;
            }

            // Excel のフォント名は大文字小文字を区別しない。
            var name = fontName.Trim().ToUpperInvariant();
            return name switch
            {
                "CALIBRI" => 7.0,
                "ＭＳ Ｐゴシック" or "MS PGOTHIC" or "游ゴシック" or "YU GOTHIC" => 8.0,
                _ => DefaultMaxDigitWidthPx,
            };
        }
    }
}
