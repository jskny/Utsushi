using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.ReportDefinitions.Model;

/// <summary>
/// 帳票定義と結び付いた1帳票分の内部モデル。置換後もこの型のまま(値のみ更新)受け渡される。
/// </summary>
/// <param name="Definition">帳票定義。</param>
/// <param name="Sheet">対象シート。</param>
/// <param name="OverflowByCell">
/// 置換を適用したセルの、はみ出し時の挙動。Layout レイヤーが参照する(要件2.5)。
/// 置換対象でないセルは含まれない。
/// </param>
public sealed record ReportModel(
    ReportDefinition Definition,
    SheetModel Sheet,
    IReadOnlyDictionary<CellAddress, OverflowBehavior> OverflowByCell)
{
    /// <summary>置換をまだ適用していない初期状態のモデルを作る。</summary>
    public static ReportModel Create(ReportDefinition definition, SheetModel sheet) =>
        new(definition, sheet, new Dictionary<CellAddress, OverflowBehavior>());

    /// <summary>指定セルのはみ出し挙動を返す。置換対象でない場合は null。</summary>
    public OverflowBehavior? GetOverflowBehavior(CellAddress address) =>
        OverflowByCell.TryGetValue(address, out var behavior) ? behavior : null;
}
