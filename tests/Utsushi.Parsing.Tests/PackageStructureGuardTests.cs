using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 監査で見つかったパッケージ構造の安全弁の穴(要件6.7)の回帰テスト。
    /// <list type="bullet">
    /// <item>展開できない <c>.rels</c> が1つあると後ろの <c>.rels</c> が検査されなかった(Corrupted にする)</item>
    /// <item><c>[Content_Types].xml</c> の大きさ・要素数の上限(<c>MaxContentTypesBytes</c>/<c>MaxContentTypesEntries</c>)</item>
    /// <item>パートのルート要素が違うと <see cref="InvalidDataException"/> が漏れた(Corrupted にする)</item>
    /// <item>ZIPエントリ数(<c>MaxZipEntries</c>)と全関係の数の合計(<c>MaxRelationshipsPerPackage</c>)の上限</item>
    /// <item>パッケージ全体の要素数の上限(<c>MaxXmlElementsPerPackage</c>)</item>
    /// </list>
    /// </summary>
    public sealed class PackageStructureGuardTests
    {
        private const string WorkbookRelsEntry = "xl/_rels/workbook.xml.rels";
        private const string ContentTypesEntry = "[Content_Types].xml";

        private readonly OpenXmlWorkbookReader _reader = new();

        // --- 壊れた .rels -------------------------------------------------------------

        [Fact]
        public void 展開できない関係パートはCorruptedになり後ろの関係パートの検査も飛ばさない()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                // 先頭側に壊れた .rels を置き、後ろの workbook.xml.rels には上限を超える関係を入れる。
                SafetyLimitWorkbookFixtures.AddExternalRelationships(path, WorkbookRelsEntry, OpenXmlWorkbookReader.MaxRelationshipsPerPart + 1);
                AddEntryFirst(path, "_rels/aaa.xml.rels", new string(' ', 4096));
                SafetyLimitWorkbookFixtures.CorruptEntryCompressedData(path, "_rels/aaa.xml.rels");

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
                Assert.Contains("_rels/aaa.xml.rels", ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- [Content_Types].xml ---------------------------------------------------

        [Fact]
        public void ContentTypesの要素数が上限を超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                AddOverrides(path, OpenXmlWorkbookReader.MaxContentTypesEntries);

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Contains(ContentTypesEntry, ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData("[CONTENT_TYPES].XML")]
        [InlineData("[Content_Type\u017F].xml")]
        public void 大文字にするとContentTypesと一致する名前のエントリも検査する(string entryName)
        {
            // System.IO.Packaging は ToUpperInvariant した名前で [Content_Types].xml を探す。ſ(U+017F)は大文字にすると S になる。
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                AddOverrides(path, OpenXmlWorkbookReader.MaxContentTypesEntries);
                RenameEntry(path, ContentTypesEntry, entryName);

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Theory]
        [InlineData(3, 3, false)]
        [InlineData(4, 3, true)]
        public void ContentTypesの要素数は上限ちょうどまで通る(int elements, int max, bool throws)
        {
            using var xml = ContentTypesXml(elements);

            void Act() => OpenXmlWorkbookReader.EnsureContentTypesWithinLimits(xml, 1024 * 1024, max, ContentTypesEntry, null);

            if (throws)
            {
                Assert.Equal(InvalidExcelFileReason.TooLarge, Assert.Throws<InvalidExcelFileException>(Act).Reason);
            }
            else
            {
                Act();
            }
        }

        [Fact]
        public void ContentTypesの大きさが上限を超えるとTooLargeになる()
        {
            using var xml = ContentTypesXml(100);

            var ex = Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureContentTypesWithinLimits(xml, 1024, 10_000, ContentTypesEntry, "report"));

            Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
            Assert.Contains("大きさ", ex.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ContentTypesの深い入れ子はTooLargeになる()
        {
            const int depth = 1000;
            var text = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + string.Concat(Enumerable.Repeat("<Q>", depth)) + string.Concat(Enumerable.Repeat("</Q>", depth)) + "</Types>";
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(text));

            Assert.Throws<InvalidExcelFileException>(
                () => OpenXmlWorkbookReader.EnsureContentTypesWithinLimits(xml, 1024 * 1024, 10_000, ContentTypesEntry, null));
        }

        // --- ルート要素の不一致 -------------------------------------------------------

        [Fact]
        public void ワークブックのルート要素が違うとCorruptedになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path, SafetyLimitWorkbookFixtures.EntryNameEndingWith(path, "workbook.xml"), "<?xml version=\"1.0\"?><foo/>");

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ワークシートのルート要素が違うとCorruptedになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                SafetyLimitWorkbookFixtures.ReplaceEntry(
                    path,
                    SafetyLimitWorkbookFixtures.FirstWorksheetEntryName(path),
                    "<?xml version=\"1.0\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");

                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.ReadFile(path));
                Assert.Equal(InvalidExcelFileReason.Corrupted, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- ZIPエントリ数・関係の数の合計 -------------------------------------------------

        [Fact]
        public void ZIPエントリ数が上限を超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                using var stream = File.OpenRead(path);
                var entryCount = OpenXmlWorkbookReader.TryReadDeclaredZipEntryCount(stream);
                using (var archive = new ZipArchive(File.OpenRead(path), ZipArchiveMode.Read))
                {
                    Assert.Equal(archive.Entries.Count, entryCount);
                }

                stream.Position = 0;

                OpenXmlWorkbookReader.GuardPackageSize(stream, long.MaxValue, (int)entryCount!.Value, int.MaxValue, null);

                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => OpenXmlWorkbookReader.GuardPackageSize(stream, long.MaxValue, (int)entryCount.Value - 1, int.MaxValue, "report"));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Contains("エントリの数", ex.Message, StringComparison.Ordinal);
                Assert.Equal(0, stream.Position);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 全関係パートの関係の数の合計が上限を超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                // _rels/.rels(1件)と workbook.xml.rels(1件 + 追加分)の合計で判定する。
                SafetyLimitWorkbookFixtures.AddExternalRelationships(path, WorkbookRelsEntry, 10);
                using var stream = File.OpenRead(path);

                OpenXmlWorkbookReader.GuardPackageSize(stream, long.MaxValue, int.MaxValue, 12, null);

                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => OpenXmlWorkbookReader.GuardPackageSize(stream, long.MaxValue, int.MaxValue, 11, null));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Contains("関係パート全体", ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 既定の上限はdesignの安全弁一覧の値()
        {
            Assert.Equal(10_000, OpenXmlWorkbookReader.MaxZipEntries);
            Assert.Equal(50_000, OpenXmlWorkbookReader.MaxRelationshipsPerPackage);
            Assert.Equal(4L * 1024 * 1024, OpenXmlWorkbookReader.MaxContentTypesBytes);
            Assert.Equal(10_000, OpenXmlWorkbookReader.MaxContentTypesEntries);
            Assert.Equal(8_000_000L, OpenXmlWorkbookReader.MaxXmlElementsPerPackage);
        }

        // --- パッケージ全体の要素数 -------------------------------------------------------

        [Fact]
        public void パッケージ全体の要素の数が上限を超えるとTooLargeになる()
        {
            var path = SafetyLimitWorkbookFixtures.CreateWorkbook();
            try
            {
                using var document = SpreadsheetDocument.Open(path, isEditable: false);
                var total = document.GetAllParts()
                    .Where(p => p.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase))
                    .Sum(p =>
                    {
                        using var s = p.GetStream(FileMode.Open, FileAccess.Read);
                        return OpenXmlWorkbookReader.EnsureXmlWithinLimits(s, 256, long.MaxValue, p.Uri.ToString(), null);
                    });

                OpenXmlWorkbookReader.GuardXmlParts(document, total, null);

                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => OpenXmlWorkbookReader.GuardXmlParts(document, total - 1, "report"));
                Assert.Equal(InvalidExcelFileReason.TooLarge, ex.Reason);
                Assert.Contains("全体", ex.Message, StringComparison.Ordinal);
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- ヘルパー -------------------------------------------------------------

        private static MemoryStream ContentTypesXml(int elements)
        {
            var builder = new StringBuilder("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            for (var i = 0; i < elements; i++)
            {
                builder.Append("<Override PartName=\"/x/p").Append(i).Append(".xml\" ContentType=\"application/xml\"/>");
            }

            builder.Append("</Types>");
            return new MemoryStream(Encoding.UTF8.GetBytes(builder.ToString()));
        }

        /// <summary>既存の [Content_Types].xml に Override を <paramref name="count"/> 件足す(既存分と合わせて上限を超える)。</summary>
        private static void AddOverrides(string path, int count) =>
            SafetyLimitWorkbookFixtures.ModifyEntry(path, ContentTypesEntry, xml =>
            {
                var builder = new StringBuilder(count * 80);
                for (var i = 0; i < count; i++)
                {
                    builder.Append("<Override PartName=\"/x/p").Append(i).Append(".xml\" ContentType=\"application/xml\"/>");
                }

                var end = xml.LastIndexOf("</Types>", StringComparison.Ordinal);
                return xml.Substring(0, end) + builder + xml.Substring(end);
            });

        /// <summary>ZIP のエントリの名前を変える。</summary>
        private static void RenameEntry(string path, string from, string to)
        {
            using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
            var source = archive.GetEntry(from)!;
            byte[] content;
            using (var stream = source.Open())
            using (var buffer = new MemoryStream())
            {
                stream.CopyTo(buffer);
                content = buffer.ToArray();
            }

            source.Delete();
            using var destination = archive.CreateEntry(to).Open();
            destination.Write(content, 0, content.Length);
        }

        /// <summary>ZIP の先頭にエントリを追加する(既存のエントリをすべて後ろへ書き直す)。</summary>
        private static void AddEntryFirst(string path, string entryName, string content)
        {
            var copy = path + ".tmp";
            using (var source = ZipFile.OpenRead(path))
            using (var destination = ZipFile.Open(copy, ZipArchiveMode.Create))
            {
                using (var stream = destination.CreateEntry(entryName, CompressionLevel.Optimal).Open())
                {
                    var bytes = Encoding.UTF8.GetBytes(content);
                    stream.Write(bytes, 0, bytes.Length);
                }

                foreach (var entry in source.Entries)
                {
                    using var from = entry.Open();
                    using var to = destination.CreateEntry(entry.FullName).Open();
                    from.CopyTo(to);
                }
            }

            File.Delete(path);
            File.Move(copy, path);
        }
    }
}
