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

        public void Add(FitIssue issue)
        {
            if (_issues.Count < MaxIssues && _seen.Add((issue.Cell, issue.Kind)))
            {
                _issues.Add(issue);
            }
        }
    }
}
