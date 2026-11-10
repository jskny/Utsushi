using System.Collections.Generic;
using Utsushi.Core.Exceptions;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Substitution
{
    /// <summary>
    /// セル番地を指定して、セルの値を書き換える(要件2.7、14.2・14.3)。書式には触れない。
    /// </summary>
    public interface ICellSubstitutor
    {
        /// <summary>
        /// セル番地(A1形式)を指定して文字列で上書きした新しい <see cref="ReportModel"/> を返す。入力のモデルは変更しない(要件2.7)。
        /// </summary>
        /// <remarks>
        /// テンプレートに無いセルも指定できる。はみ出し時の挙動はセルの書式に従う。
        /// </remarks>
        /// <param name="report">上書き前の帳票モデル。</param>
        /// <param name="cellOverrides">セル番地(A1形式。例: <c>"B5"</c>) → 上書き後の文字列。</param>
        /// <exception cref="InvalidCellOverrideAddressException">
        /// セル番地がA1形式として解釈できない場合(要件2.8)。
        /// </exception>
        /// <exception cref="NonAnchorMergedCellOverrideException">
        /// 指定したセルが結合セル範囲内にあり、かつ先頭(アンカー)セルではない場合(要件2.9)。
        /// Layoutレイヤーは結合範囲のアンカーセルの値しか描画しないため、アンカー以外を
        /// 指定すると値がモデルには反映されてもPDFには一切出力されない静かなデータ欠落になる。
        /// </exception>
        ReportModel ApplyCellOverrides(ReportModel report, IReadOnlyDictionary<string, string> cellOverrides);

        /// <summary>
        /// セル番地(A1形式)を直接指定して、数値で上書きした新しい <see cref="ReportModel"/> を返す(要件14.3)。
        /// 対象のセルは数値のセルになり、セルの数値書式(要件4.8)と色の指定(要件4.12)で表示される。
        /// 番地の検証・例外は <see cref="ApplyCellOverrides"/> と同じ。件数の上限は <see cref="ApplyCellOverrides"/> で上書きした
        /// セルと合わせて数える。<see cref="ApplyCellOverrides"/> と同じセルを指定した場合は、後から適用したこちらの値になる。
        /// </summary>
        /// <param name="report">上書き前の帳票モデル。</param>
        /// <param name="numericOverrides">セル番地(A1形式) → 数値。NaN・無限大は不可。</param>
        /// <exception cref="InvalidSubstitutionValueException">数値が NaN・無限大の場合(要件14.4)。</exception>
        ReportModel ApplyNumericCellOverrides(ReportModel report, IReadOnlyDictionary<string, double> numericOverrides);
    }
}
