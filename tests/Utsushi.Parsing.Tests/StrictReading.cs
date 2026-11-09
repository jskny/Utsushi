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
                    workbook.UnsupportedElements, workbook.UnsupportedElementsTruncated, options.ReportCode);
            }

            return workbook;
        }

        /// <summary>例外に、指定の種類(と補足の情報の先頭)の未対応要素が含まれることを確かめる。</summary>
        public static UnsupportedElement AssertContains(
            UnsupportedWorkbookElementException ex, UnsupportedElementKind kind, string? detailPrefix = null)
        {
            var match = ex.Elements.FirstOrDefault(
                e => e.Kind == kind && (detailPrefix is null || (e.Detail ?? string.Empty).StartsWith(detailPrefix, StringComparison.Ordinal)));
            Assert.True(match is not null, $"未対応要素 {kind} {detailPrefix} が含まれていません: {string.Join(" / ", ex.Elements)}");
            return match!;
        }
    }
}
