using System;
using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.ReportDefinitions.Model
{
    /// <summary>
    /// 置換後の文字列がセル幅に収まらない場合の挙動(要件2.5)。
    /// </summary>
    /// <remarks>実際の折り返し/縮小計算は Layout レイヤーが行い、ここでは指定値を保持するだけである。</remarks>
    public enum OverflowBehavior
    {
        /// <summary>セル幅に収まるようフォントを縮小する(Excel の「縮小して全体を表示する」相当)。</summary>
        Shrink = 0,

        /// <summary>セル幅で切り詰める。</summary>
        Clip,

        /// <summary>セル内で折り返す。</summary>
        Wrap,

        /// <summary>はみ出しを許容し、隣接セルへ描画する(Excel の既定動作相当)。</summary>
        Overflow,
    }

    /// <summary>サポート外要素を検出したときの挙動(要件1.5)。</summary>
    public enum UnsupportedElementPolicy
    {
        /// <summary>無視して変換を続行する。</summary>
        Ignore = 0,

        /// <summary>エラーとして変換を中止する。</summary>
        Error,
    }

    /// <summary>
    /// 置換対象フィールド1件の定義。
    /// </summary>
    /// <param name="Key">置換キー(論理名)。セル番地ではなくこのキーで指定する。</param>
    /// <param name="Cell">置換先セル番地。</param>
    /// <param name="Required">必須かどうか。true の場合、値が渡されないとエラーになる(要件2.4)。</param>
    /// <param name="Overflow">
    /// セル幅に収まらない場合の挙動(要件2.5)。
    /// <c>null</c> は帳票定義で未指定であることを表し、この場合は Excel 側のセル書式
    /// (<c>wrapText</c> / <c>shrinkToFit</c>)に従う。
    /// </param>
    public sealed record SubstitutionFieldDefinition(
        string Key,
        CellAddress Cell,
        bool Required,
        OverflowBehavior? Overflow);

    /// <summary>
    /// 1帳票テンプレートに対応する帳票定義。
    /// </summary>
    /// <remarks>
    /// 帳票ごとの差異はすべてこの型(とその元になるJSON)に閉じ込める。
    /// 共通レイヤーのコードに帳票固有の分岐を持ち込まないこと(`.kiro/steering/structure.md`)。
    /// </remarks>
    /// <param name="ReportCode">帳票コード。定義ディレクトリ名と一致させる。</param>
    /// <param name="SheetName">対象シート名。</param>
    /// <param name="SubstitutionFields">置換対象フィールド。</param>
    /// <param name="ToleranceMm">Excel表示との許容誤差(mm)。ゴールデンテストの比較閾値に用いる(要件4.5)。</param>
    /// <param name="UnsupportedElements">サポート外要素の扱い(要件1.5)。</param>
    /// <param name="MaxDigitWidthPx">
    /// 列幅(文字数単位)→ピクセル換算に用いる標準フォントの最大数字幅(96dpiのピクセル)。
    /// Excel の換算式が標準フォント依存であるため、帳票ごとに定義側で指定できるようにしている。
    /// </param>
    /// <param name="PrintAreaOverride">
    /// 帳票定義側で印刷範囲を上書きする場合の範囲。null ならExcelの印刷範囲設定に従う。
    /// </param>
    public sealed record ReportDefinition(
        string ReportCode,
        string SheetName,
        IReadOnlyList<SubstitutionFieldDefinition> SubstitutionFields,
        double ToleranceMm,
        UnsupportedElementPolicy UnsupportedElements,
        double MaxDigitWidthPx,
        CellRange? PrintAreaOverride)
    {
        /// <summary>Excel の標準フォント(11pt)における最大数字幅の既定値(ピクセル)。</summary>
        public const double DefaultMaxDigitWidthPx = 7.0;

        /// <summary>許容誤差の既定値(mm)。</summary>
        public const double DefaultToleranceMm = 0.5;

        private Dictionary<string, SubstitutionFieldDefinition>? _fieldsByKey;

        /// <summary>
        /// 帳票定義なしモード(要件12)で使う、既定値だけからなる帳票定義を合成する。
        /// 置換対象フィールドは持たず、サポート外要素は無視し、印刷範囲はExcelの設定に従う。
        /// </summary>
        /// <param name="documentName">
        /// 帳票コードの代わりに使う文書名(PDFタイトル・ヘッダー/フッターの <c>&amp;F</c>・エラー情報に使われる。要件12.5)。
        /// </param>
        /// <param name="sheetName">変換対象のシート名。</param>
        /// <param name="maxDigitWidthPx">
        /// 列幅の換算に使う最大数字幅(ピクセル)。通常は <see cref="EstimateMaxDigitWidthPx"/> でブックの標準フォントから求める。
        /// </param>
        public static ReportDefinition CreateWithoutDefinition(
            string documentName, string sheetName, double maxDigitWidthPx = DefaultMaxDigitWidthPx) =>
            new(
                documentName ?? throw new ArgumentNullException(nameof(documentName)),
                sheetName ?? throw new ArgumentNullException(nameof(sheetName)),
                Array.Empty<SubstitutionFieldDefinition>(),
                DefaultToleranceMm,
                UnsupportedElementPolicy.Ignore,
                maxDigitWidthPx > 0 && !double.IsNaN(maxDigitWidthPx) && !double.IsInfinity(maxDigitWidthPx)
                    ? maxDigitWidthPx
                    : throw new ArgumentOutOfRangeException(nameof(maxDigitWidthPx), maxDigitWidthPx, "最大数字幅は正の数である必要があります。"),
                PrintAreaOverride: null);

        /// <summary>
        /// ブックの標準フォント(名前・サイズ)から、Excelが列幅の換算に使う最大数字幅(96dpiのピクセル)を見積もる
        /// (帳票定義なしモード、要件12.3)。
        /// </summary>
        /// <remarks>
        /// Excelの既定の標準フォントについて、既定の列幅の表示(英語版 Calibri 11 の「8.43(64ピクセル)」、日本語版
        /// ＭＳ Ｐゴシック/游ゴシック 11 の「8.38(72ピクセル)」)と列幅の換算式から逆算した値を持つ。表に無いフォント・サイズは
        /// <see cref="DefaultMaxDigitWidthPx"/> を返す。帳票定義ありの変換では使わない(定義の <c>maxDigitWidthPx</c> を使う)。
        /// </remarks>
        public static double EstimateMaxDigitWidthPx(string? fontName, double sizePt)
        {
            if (Math.Abs(sizePt - 11.0) > 0.01 || fontName is null)
            {
                return DefaultMaxDigitWidthPx;
            }

            return fontName.Trim() switch
            {
                "Calibri" => 7.0,
                "ＭＳ Ｐゴシック" or "MS PGothic" or "游ゴシック" or "Yu Gothic" => 8.0,
                _ => DefaultMaxDigitWidthPx,
            };
        }

        /// <summary>置換キーからフィールド定義を引く。</summary>
        public bool TryGetField(string key, out SubstitutionFieldDefinition field)
        {
            _fieldsByKey ??= BuildIndex(SubstitutionFields);
            return _fieldsByKey.TryGetValue(key, out field!);
        }

        /// <summary>必須の置換キーを列挙する。</summary>
        public IEnumerable<SubstitutionFieldDefinition> RequiredFields
        {
            get
            {
                foreach (var field in SubstitutionFields)
                {
                    if (field.Required)
                    {
                        yield return field;
                    }
                }
            }
        }

        private static Dictionary<string, SubstitutionFieldDefinition> BuildIndex(
            IReadOnlyList<SubstitutionFieldDefinition> fields)
        {
            var index = new Dictionary<string, SubstitutionFieldDefinition>(StringComparer.Ordinal);
            foreach (var field in fields)
            {
                index[field.Key] = field;
            }

            return index;
        }
    }
}
