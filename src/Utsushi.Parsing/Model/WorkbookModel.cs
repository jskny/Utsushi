using System;
using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>
    /// ワークブック全体の内部モデル。Parsing レイヤーが上位レイヤーへ渡す唯一の型。
    /// </summary>
    /// <param name="Sheets">シート(ブック内の並び順)。</param>
    /// <param name="DefaultFont">ブックの標準フォント。列幅のポイント換算に用いる。</param>
    public sealed record WorkbookModel(IReadOnlyList<SheetModel> Sheets, FontStyle DefaultFont)
    {
        /// <summary>
        /// 読み込んだシート(と、外部参照のようなブック全体)で見つかった未対応要素(excel2pdf-consolidation 要件2)。
        /// 読み込みは未対応要素を出力せずに読み飛ばしたうえで、ここに記録する。
        /// </summary>
        public IReadOnlyList<UnsupportedElement> UnsupportedElements { get; init; } = Array.Empty<UnsupportedElement>();

        /// <summary>未対応要素の件数の上限に達し、記録しなかった要素があるかどうか。</summary>
        public bool UnsupportedElementsTruncated { get; init; }

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
