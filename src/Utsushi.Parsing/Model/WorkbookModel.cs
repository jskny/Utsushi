using System;
using System.Collections.Generic;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// ワークブック全体の内部モデル。Parsing レイヤーが上位レイヤーへ渡す唯一の型。
    /// </summary>
    /// <param name="Sheets">シート(ブック内の並び順)。</param>
    /// <param name="DefaultFont">ブックの標準フォント。列幅のポイント換算に用いる。</param>
    public sealed record WorkbookModel(IReadOnlyList<SheetModel> Sheets, FontStyle DefaultFont)
    {
        /// <summary>シート名でシートを取得する。見つからない場合は null。</summary>
        public SheetModel? FindSheet(string name)
        {
            for (var i = 0; i < Sheets.Count; i++)
            {
                if (string.Equals(Sheets[i].Name, name, StringComparison.Ordinal))
                {
                    return Sheets[i];
                }
            }

            return null;
        }
    }
}
