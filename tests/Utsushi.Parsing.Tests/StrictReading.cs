using System;
using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 未対応要素を例外にする読み込み(以前の <c>UnsupportedElementBehavior.Error</c> 相当)。
    /// 読み込みは常に記録だけを行うため(excel2pdf-consolidation design.md 2.2)、エラーにするのは呼び出し元(ファサード)と同じく、ここで行う。
    /// </summary>
    internal static class StrictReading
    {
        public static WorkbookModel ReadStrict(this OpenXmlWorkbookReader reader, Stream stream, WorkbookReadOptions? options = null)
        {
            options ??= WorkbookReadOptions.Default;
            var workbook = reader.Read(stream, options);
            if (workbook.UnsupportedElements.Count > 0)
            {
                throw new UnsupportedWorkbookElementException(
                    workbook.UnsupportedElements, workbook.UnsupportedElementsTruncated, options.DocumentName);
            }

            return workbook;
        }

        /// <summary>例外に、指定の種類(と補足の情報の先頭)の未対応要素がちょうど1件含まれることを確かめる。</summary>
        public static UnsupportedElement AssertContains(
            UnsupportedWorkbookElementException ex, UnsupportedElementKind kind, string? detailPrefix = null)
        {
            // 1件だけであることも確かめる(同じ要素の二重の記録を見逃さないため)。
            var matches = ex.Elements
                .Where(e => e.Kind == kind && (detailPrefix is null || (e.Detail ?? string.Empty).StartsWith(detailPrefix, StringComparison.Ordinal)))
                .ToList();
            Assert.True(matches.Count == 1, $"未対応要素 {kind} {detailPrefix} が1件ではありません({matches.Count}件): {string.Join(" / ", ex.Elements)}");
            return matches[0];
        }
    }
}
