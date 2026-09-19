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
}
