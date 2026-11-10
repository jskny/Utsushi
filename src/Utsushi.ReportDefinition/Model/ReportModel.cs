using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.ReportDefinitions.Model
{
    /// <summary>
    /// 変換の設定と対象シートを組にした1文書分の内部モデル。値の書き込み後もこの型のまま(値のみ更新)受け渡される。
    /// </summary>
    /// <param name="Settings">変換の設定。</param>
    /// <param name="Sheet">対象シート。</param>
    /// <param name="DefaultFont">
    /// ブックの標準フォント。セル書式を持たない描画(ヘッダー/フッター)の既定として使う。
    /// </param>
    public sealed record ReportModel(
        ConversionSettings Settings,
        SheetModel Sheet,
        FontStyle DefaultFont)
    {
        /// <summary>値をまだ書き込んでいない初期状態のモデルを作る。</summary>
        public static ReportModel Create(ConversionSettings settings, SheetModel sheet, FontStyle? defaultFont = null) =>
            new(settings, sheet, defaultFont ?? FontStyle.Default);

        /// <summary>
        /// 呼び出し元から空でない値が書き込まれたセル(<c>SetText</c>・<c>SetValue</c>)。
        /// </summary>
        /// <remarks>
        /// テンプレート自身の文字列と違い、差し込み値はExcel上で人が目視確認していない。
        /// Layout レイヤーはこの集合に含まれるセルに限り、印刷範囲外(要件2.13)・
        /// 折り返し行の欠落(要件2.14)をエラーとして検出する。
        /// </remarks>
        public IReadOnlySet<CellAddress> SubstitutedCells { get; init; } = new HashSet<CellAddress>();

        /// <summary>
        /// セル番地直接指定(要件2.7)で値を上書きしたセル(空文字の書き込みを含む)。書き込みの件数の上限の判定に使う。
        /// </summary>
        public IReadOnlySet<CellAddress> OverriddenCells { get; init; } = new HashSet<CellAddress>();
    }
}
