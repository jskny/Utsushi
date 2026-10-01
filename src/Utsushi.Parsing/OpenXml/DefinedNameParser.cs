using System;
using System.Collections.Generic;
using System.Globalization;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// <c>definedNames</c> の <c>_xlnm.Print_Area</c> / <c>_xlnm.Print_Titles</c> を解釈する。
    /// </summary>
    internal static class DefinedNameParser
    {
        /// <summary>印刷範囲の組み込み定義名。</summary>
        public const string PrintAreaName = "_xlnm.Print_Area";

        /// <summary>印刷タイトルの組み込み定義名。</summary>
        public const string PrintTitlesName = "_xlnm.Print_Titles";

        /// <summary>
        /// 印刷範囲の定義("Sheet1!$A$1:$H$40" 形式、カンマ区切りで複数可)を解釈する。
        /// </summary>
        public static IReadOnlyList<CellRange> ParsePrintArea(string? definition) => ParsePrintArea(definition, int.MaxValue);

        /// <summary>
        /// <see cref="ParsePrintArea(string?)"/> と同じだが、<paramref name="maxCount"/> + 1 個まで読んだ時点で打ち切る
        /// (要件6.8)。戻り値の個数が <paramref name="maxCount"/> を超えていれば、上限を超えていたことを表す。
        /// 上限を大きく超える定義でも、全件をリストにしてから比べることはしない(security-reviewer指摘)。
        /// </summary>
        public static IReadOnlyList<CellRange> ParsePrintArea(string? definition, int maxCount)
        {
            var result = new List<CellRange>();
            foreach (var reference in SplitReferences(definition))
            {
                var range = StripSheetName(reference);
                if (CellRange.TryParse(range, out var parsed))
                {
                    result.Add(parsed);
                    if (result.Count > maxCount)
                    {
                        break;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 印刷タイトルの定義("Sheet1!$1:$3,Sheet1!$A:$B" 形式)を解釈する。
        /// </summary>
        public static PrintTitles ParsePrintTitles(string? definition)
        {
            int? firstRow = null, lastRow = null, firstColumn = null, lastColumn = null;

            foreach (var reference in SplitReferences(definition))
            {
                var range = StripSheetName(reference).Replace("$", string.Empty);
                var parts = range.Split(':');
                if (parts.Length != 2)
                {
                    continue;
                }

                if (int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var rowStart)
                    && int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var rowEnd))
                {
                    // 行番号の範囲外($0 や最大行を超える値)の指定は、行の印刷タイトルなしとして無視する
                    // (そのまま渡すと Layout で ArgumentOutOfRangeException になっていた)。
                    if (IsValidRow(rowStart) && IsValidRow(rowEnd))
                    {
                        firstRow = Math.Min(rowStart, rowEnd);
                        lastRow = Math.Max(rowStart, rowEnd);
                    }

                    continue;
                }

                if (CellAddress.TryParseColumnName(parts[0], out var colStart)
                    && CellAddress.TryParseColumnName(parts[1], out var colEnd))
                {
                    firstColumn = Math.Min(colStart, colEnd);
                    lastColumn = Math.Max(colStart, colEnd);
                }
            }

            return new PrintTitles(firstRow, lastRow, firstColumn, lastColumn);
        }

        private static bool IsValidRow(int row) => row >= 1 && row <= CellAddress.MaxRow;

        /// <summary>シート名(引用符付きを含む)を除いた参照部分を返す。</summary>
        private static string StripSheetName(string reference)
        {
            var index = reference.LastIndexOf('!');
            return index < 0 ? reference : reference.Substring(index + 1);
        }

        /// <summary>引用符付きシート名の中のカンマを壊さずに参照を分割する。</summary>
        private static IEnumerable<string> SplitReferences(string? definition)
        {
            if (string.IsNullOrWhiteSpace(definition))
            {
                yield break;
            }

            var text = definition!;
            var start = 0;
            var inQuotes = false;

            for (var i = 0; i < text.Length; i++)
            {
                var c = text[i];
                if (c == '\'')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    var part = text.Substring(start, i - start).Trim();
                    if (part.Length > 0)
                    {
                        yield return part;
                    }

                    start = i + 1;
                }
            }

            var last = text.Substring(start).Trim();
            if (last.Length > 0)
            {
                yield return last;
            }
        }
    }
}
