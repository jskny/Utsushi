using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Layout.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// 文字の収まりの確認(要件13)の結果を、全ページを通して集める。同じセル・同じ種類は最初に現れたページの1件にまとめる
    /// (印刷タイトルのように複数ページに現れるセルを、ページごとに数えない。要件13.4)。
    /// </summary>
    internal sealed class FitIssueCollector
    {
        /// <summary>
        /// 集める件数の上限。セルの数に比例して増えるため、列幅の足りない大きな一覧表などで結果が際限なく大きくならないようにする。
        /// ログに残す用途では、これを超えて列挙する意味は小さい。
        /// </summary>
        internal const int MaxIssues = 10_000;

        private readonly List<FitIssue> _issues = new();
        private readonly HashSet<(CellAddress Cell, FitIssueKind Kind)> _seen = new();

        public IReadOnlyList<FitIssue> Issues => _issues;

        /// <summary>上限に達し、記録しなかった箇所があるかどうか。</summary>
        public bool IsTruncated { get; private set; }

        /// <summary>
        /// この(セル, 種類)を記録するかどうか。既に記録済みなら false。上限に達していれば <see cref="IsTruncated"/> を立てて false。
        /// 記録しないものの説明の文を組み立てずに済むよう、<see cref="Add"/> の前に呼ぶ。
        /// </summary>
        public bool Accepts(CellAddress cell, FitIssueKind kind)
        {
            if (_seen.Contains((cell, kind)))
            {
                return false;
            }

            if (_issues.Count >= MaxIssues)
            {
                IsTruncated = true;
                return false;
            }

            return true;
        }

        public void Add(FitIssue issue)
        {
            if (Accepts(issue.Cell, issue.Kind))
            {
                _seen.Add((issue.Cell, issue.Kind));
                _issues.Add(issue);
            }
        }
    }
}
