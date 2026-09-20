using System.Collections.Generic;
using Utsushi.Core.Exceptions;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Substitution;

/// <summary>
/// 置換キーと値の辞書を受け取り、帳票定義に登録された対象セルの値のみを書き換える。
/// </summary>
public interface ICellSubstitutor
{
    /// <summary>
    /// 置換を適用した新しい <see cref="ReportModel"/> を返す。入力のモデルは変更しない。
    /// </summary>
    /// <param name="report">置換前の帳票モデル。</param>
    /// <param name="values">置換キー → 置換後の文字列。</param>
    /// <exception cref="SubstitutionKeyNotFoundException">
    /// 帳票定義に存在しない置換キーが指定された場合(要件2.3)。
    /// </exception>
    /// <exception cref="RequiredSubstitutionValueMissingException">
    /// 必須の置換キーに対応する値が渡されていない場合(要件2.4)。
    /// </exception>
    ReportModel Apply(ReportModel report, IReadOnlyDictionary<string, string> values);

    /// <summary>
    /// 帳票定義の置換キーを経由せず、セル番地(A1形式)を直接指定して値を上書きした
    /// 新しい <see cref="ReportModel"/> を返す。入力のモデルは変更しない(要件2.7)。
    /// </summary>
    /// <remarks>
    /// <see cref="Apply"/> と異なり、未知キー検証(要件2.3)・必須キー検証(要件2.4)の対象外であり、
    /// 帳票定義に登録の無いセルも指定できる。はみ出し時の挙動もExcel側のセル書式に従う(常に未指定扱い)。
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
}
