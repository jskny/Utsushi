using System.Collections.Generic;
using System.Globalization;
using Utsushi.Core;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// 1回の読み込みで見つかった未対応要素を集める(excel2pdf-consolidation 要件2.4・2.5、design.md 2.2)。
    /// </summary>
    /// <remarks>
    /// 記録の単位(要件2.5):
    /// <list type="bullet">
    /// <item>枚数・個数の上限(<see cref="AddCountLimitExceeded"/>)は、シートと上限の種類ごとに1件。超えた後の要素を1つずつ記録すると上限が意味をなさなくなるため。</item>
    /// <item>それ以外(<see cref="Add"/>)は要素ごとに1件。</item>
    /// </list>
    /// 記録の件数は <see cref="MaxElements"/> で打ち切り、<see cref="IsTruncated"/> を立てる。
    /// 補足の情報(<c>detail</c>)はファイル由来の任意の文字列になりうるため、<see cref="DiagnosticText.Normalize"/> で正規化する。
    /// </remarks>
    internal sealed class UnsupportedElementCollector
    {
        /// <summary>記録する件数の上限。</summary>
        internal const int MaxElements = 10_000;

        private readonly int _maxElements;
        private readonly List<UnsupportedElement> _elements = new();
        private readonly HashSet<(string? SheetName, string LimitName)> _countLimits = new();

        public UnsupportedElementCollector()
            : this(MaxElements)
        {
        }

        /// <summary>テスト用: 件数の上限を指定して作る(<c>.kiro/steering/tech.md</c>「テスト戦略」)。</summary>
        internal UnsupportedElementCollector(int maxElements)
        {
            _maxElements = maxElements;
        }

        /// <summary>記録した未対応要素(記録した順)。</summary>
        public IReadOnlyList<UnsupportedElement> Elements => _elements;

        /// <summary>件数の上限に達し、記録しなかった要素があるかどうか。</summary>
        public bool IsTruncated { get; private set; }

        /// <summary>未対応要素を1件記録する。</summary>
        /// <param name="message">説明の文。セル・図形の文字列を含めないこと(呼び出し元の責務)。</param>
        public void Add(UnsupportedElementKind kind, string? sheetName, CellAddress? anchor, string? detail, string message)
        {
            if (_elements.Count >= _maxElements)
            {
                IsTruncated = true;
                return;
            }

            _elements.Add(new UnsupportedElement(kind, sheetName, anchor, DiagnosticText.Normalize(detail), message));
        }

        /// <summary>
        /// 1つの要素の上限(画像1枚の大きさ・図形の文字数・入れ子の段数など)の超過を、要素ごとに記録する。
        /// </summary>
        public void AddElementLimitExceeded(string? sheetName, CellAddress? anchor, string limitName, long limit, string message) =>
            Add(UnsupportedElementKind.LimitExceeded, sheetName, anchor, FormatLimit(limitName, limit), message);

        /// <summary>
        /// 枚数・個数の上限(画像の枚数・図形の数・結合セル範囲の数)の超過を、シートと上限の種類ごとに1件だけ記録する。
        /// </summary>
        public void AddCountLimitExceeded(string? sheetName, CellAddress? anchor, string limitName, long limit, string message)
        {
            if (!_countLimits.Add((sheetName, limitName)))
            {
                return;
            }

            Add(UnsupportedElementKind.LimitExceeded, sheetName, anchor, FormatLimit(limitName, limit), message);
        }

        /// <summary>
        /// 描画オブジェクトを出力しない理由(<see cref="DrawingRejection"/>)を、指定のアンカーで記録する。
        /// 枚数・個数の上限はシートと上限の種類ごとに1件、それ以外は1件ずつ記録する。
        /// </summary>
        public void Add(DrawingRejection rejection, string? sheetName, CellAddress? anchor)
        {
            if (rejection.LimitName is { } limitName)
            {
                if (rejection.IsCountLimit)
                {
                    AddCountLimitExceeded(sheetName, anchor, limitName, rejection.Limit, rejection.Message);
                }
                else
                {
                    AddElementLimitExceeded(sheetName, anchor, limitName, rejection.Limit, rejection.Message);
                }

                return;
            }

            Add(rejection.Kind, sheetName, anchor, rejection.Detail, rejection.Message);
        }

        private static string FormatLimit(string limitName, long limit) =>
            limitName + ": 上限 " + limit.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 描画オブジェクト(画像・図形・接続線・グループの子要素)を出力しない理由。
    /// 読み込みの関数はこれを返し、記録は呼び出し元が行う(グループの中の要素は、グループの側でまとめて1件にするため。design.md 9)。
    /// </summary>
    /// <param name="Kind">未対応要素の種類。</param>
    /// <param name="Detail">補足の情報(上限の超過では使わない)。</param>
    /// <param name="Message">説明の文。セル・図形の文字列を含めない。</param>
    /// <param name="LimitName">上限の超過の場合、上限の種類(<c>TooManyShapes</c> など)。それ以外は <c>null</c>。</param>
    /// <param name="Limit">上限の値。</param>
    /// <param name="IsCountLimit">枚数・個数の上限(シートと上限の種類ごとに1件にまとめる)かどうか。</param>
    internal sealed record DrawingRejection(
        UnsupportedElementKind Kind,
        string? Detail,
        string Message,
        string? LimitName = null,
        long Limit = 0,
        bool IsCountLimit = false)
    {
        public static DrawingRejection CountLimit(string limitName, long limit, string message) =>
            new(UnsupportedElementKind.LimitExceeded, null, message, limitName, limit, IsCountLimit: true);

        public static DrawingRejection ElementLimit(string limitName, long limit, string message) =>
            new(UnsupportedElementKind.LimitExceeded, null, message, limitName, limit, IsCountLimit: false);
    }
}
