using System.Collections.Generic;
using Utsushi.Core;

namespace Utsushi.Layout
{
    /// <summary>
    /// 書式の近似(excel2pdf-consolidation 要件7)を、全ページを通して集める。
    /// </summary>
    /// <remarks>
    /// 重複の判定のキーは(種類, 位置, 要素の識別)とする。印刷タイトルのように複数のページに現れるものは1件にまとまり、
    /// 同じ位置にある別の要素(図形のIDなどが異なるもの)は別の件になる(要件7.4)。
    /// 件数は種類ごとに <see cref="MaxPerKind"/> で打ち切る。行全体・列全体の網かけのように件数が多くなりやすい種類で、
    /// ほかの種類が記録されなくなるのを防ぐため(design.md 3.5)。
    /// </remarks>
    internal sealed class FormatApproximationCollector
    {
        /// <summary>種類ごとの件数の上限。</summary>
        internal const int MaxPerKind = 1_000;

        private readonly int _maxPerKind;
        private readonly List<FormatApproximation> _items = new();
        private readonly HashSet<(FormatApproximationKind Kind, CellRange? Location, string? Identity)> _seen = new();
        private readonly Dictionary<FormatApproximationKind, int> _counts = new();

        public FormatApproximationCollector()
            : this(MaxPerKind)
        {
        }

        /// <summary>テスト用: 種類ごとの件数の上限を指定して作る。</summary>
        internal FormatApproximationCollector(int maxPerKind)
        {
            _maxPerKind = maxPerKind;
        }

        public IReadOnlyList<FormatApproximation> Items => _items;

        /// <summary>いずれかの種類が件数の上限に達し、記録しなかったものがあるかどうか。</summary>
        public bool IsTruncated { get; private set; }

        /// <summary>
        /// 書式の近似を1件記録する。同じ(種類, 位置, 要素の識別)が記録済みなら何もしない。
        /// </summary>
        /// <param name="identity">要素の識別(図形のID・ヘッダー/フッターの位置など)。セルの書式では <c>null</c>。</param>
        /// <param name="detail">書式の値。ファイル由来の値は <see cref="DiagnosticText.Normalize"/> で正規化して入れる。</param>
        /// <param name="message">説明の文。セル・図形の文字列を含めないこと(呼び出し元の責務)。</param>
        public void Add(
            FormatApproximationKind kind, string sheetName, CellRange? location, string? identity, string? detail, string message)
        {
            if (_seen.Contains((kind, location, identity)))
            {
                return;
            }

            _counts.TryGetValue(kind, out var count);
            if (count >= _maxPerKind)
            {
                IsTruncated = true;
                return;
            }

            _seen.Add((kind, location, identity));
            _counts[kind] = count + 1;
            _items.Add(new FormatApproximation(kind, sheetName, location, DiagnosticText.Normalize(detail), message));
        }

        /// <summary>セル1つの書式の近似を記録する。</summary>
        public void AddCell(FormatApproximationKind kind, string sheetName, CellAddress cell, string? detail, string message) =>
            Add(kind, sheetName, new CellRange(cell.Row, cell.Column, cell.Row, cell.Column), null, detail, message);
    }
}
