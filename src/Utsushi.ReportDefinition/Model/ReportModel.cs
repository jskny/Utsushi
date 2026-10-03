using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.ReportDefinitions.Model
{
    /// <summary>
    /// 帳票定義と結び付いた1帳票分の内部モデル。置換後もこの型のまま(値のみ更新)受け渡される。
    /// </summary>
    /// <param name="Definition">帳票定義。</param>
    /// <param name="Sheet">対象シート。</param>
    /// <param name="OverflowByCell">
    /// 置換を適用したセルの、はみ出し時の挙動。Layout レイヤーが参照する(要件2.5)。
    /// 置換対象でないセルは含まれない。
    /// </param>
    /// <param name="DefaultFont">
    /// ブックの標準フォント。セル書式を持たない描画(ヘッダー/フッター)の既定として使う。
    /// </param>
    public sealed record ReportModel(
        ReportDefinition Definition,
        SheetModel Sheet,
        IReadOnlyDictionary<CellAddress, OverflowBehavior> OverflowByCell,
        FontStyle DefaultFont)
    {
        /// <summary>置換をまだ適用していない初期状態のモデルを作る。</summary>
        public static ReportModel Create(ReportDefinition definition, SheetModel sheet, FontStyle? defaultFont = null) =>
            new(definition, sheet, new Dictionary<CellAddress, OverflowBehavior>(), defaultFont ?? FontStyle.Default);

        /// <summary>
        /// 呼び出し元から空でない置換値が差し込まれたセル(置換キー経由・セル番地直接指定の両方)。
        /// </summary>
        /// <remarks>
        /// テンプレート自身の文字列と違い、差し込み値はExcel上で人が目視確認していない。
        /// Layout レイヤーはこの集合に含まれるセルに限り、印刷範囲外(要件2.13)・
        /// 折り返し行の欠落(要件2.14)をエラーとして検出する。
        /// </remarks>
        public IReadOnlySet<CellAddress> SubstitutedCells { get; init; } = new HashSet<CellAddress>();

        /// <summary>
        /// セル番地直接指定(要件2.7)で値を上書きしたセル。置換キーで差し込んだ後に同じセルを上書きした場合も含む。
        /// 文字の収まりの確認(要件13.4)で、置換キーによる置換とセル番地直接指定を区別するために使う。
        /// </summary>
        public IReadOnlySet<CellAddress> OverriddenCells { get; init; } = new HashSet<CellAddress>();

        /// <summary>指定セルのはみ出し挙動を返す。置換対象でない場合は null。</summary>
        public OverflowBehavior? GetOverflowBehavior(CellAddress address) =>
            OverflowByCell.TryGetValue(address, out var behavior) ? behavior : null;
    }
}
