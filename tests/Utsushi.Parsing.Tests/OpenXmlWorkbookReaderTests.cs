using System;
using System.IO;
using System.Linq;
using System.Text;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// Excel(OOXML)読み取りの検証(要件1.1-1.5, 6.1-6.2、タスク2.1-2.5)。
    /// </summary>
    public sealed class OpenXmlWorkbookReaderTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        private SheetModel ReadInvoiceSheet()
        {
            var workbook = _reader.ReadFile(TestPaths.SampleTemplate("invoice"));
            return Assert.Single(workbook.Sheets);
        }

        [Fact]
        public void シート名とセル値を読み取る()
        {
            var sheet = ReadInvoiceSheet();

            Assert.Equal("請求書", sheet.Name);
            Assert.Equal("請求書", sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
            Assert.Equal("請求番号", sheet.GetCell(CellAddress.Parse("E3"))!.DisplayValue);
        }

        [Fact]
        public void フォント_配置_罫線_背景色を読み取る()
        {
            var sheet = ReadInvoiceSheet();

            // タイトル: 18pt 太字・中央揃え
            var title = sheet.GetCell(CellAddress.Parse("A1"))!.Style;
            Assert.Equal(18.0, title.Font.SizePt);
            Assert.True(title.Font.Bold);
            Assert.Equal(HorizontalAlignment.Center, title.HAlign);
            Assert.Equal(VerticalAlignment.Center, title.VAlign);

            // 明細見出し: 背景色つき・4辺に細線
            var header = sheet.GetCell(CellAddress.Parse("A11"))!.Style;
            Assert.False(header.BackgroundColor.IsTransparent);
            Assert.Equal(BorderLineStyle.Thin, header.Borders.Top.Style);
            Assert.Equal(BorderLineStyle.Thin, header.Borders.Bottom.Style);
            Assert.Equal(BorderLineStyle.Thin, header.Borders.Left.Style);
            Assert.Equal(BorderLineStyle.Thin, header.Borders.Right.Style);

            // ご請求金額欄: 太線の box
            var total = sheet.GetCell(CellAddress.Parse("C8"))!.Style;
            Assert.Equal(BorderLineStyle.Medium, total.Borders.Top.Style);
        }

        [Fact]
        public void 数値書式を適用した表示文字列を持つ()
        {
            var sheet = ReadInvoiceSheet();

            // 単価 320000 → "¥320,000"
            var unitPrice = sheet.GetCell(CellAddress.Parse("E12"))!;
            Assert.Equal(CellValueKind.Number, unitPrice.ValueKind);
            Assert.Equal("¥320,000", unitPrice.DisplayValue);

            // 発行日(シリアル値)→ "2026年4月1日"
            var issueDate = sheet.GetCell(CellAddress.Parse("F4"))!;
            Assert.Equal("2026年4月1日", issueDate.DisplayValue);
        }

        [Fact]
        public void 列幅と行高と結合セルを読み取る()
        {
            var sheet = ReadInvoiceSheet();

            Assert.Equal(4.0, sheet.GetColumnWidth(1));    // A列
            Assert.Equal(28.0, sheet.GetColumnWidth(2));   // B列
            Assert.Equal(30.0, sheet.GetRowHeight(1));     // 1行目

            // 定義の無い行には既定行高を返す
            Assert.Equal(13.5, sheet.GetRowHeight(500));

            var merged = sheet.FindMergedRange(CellAddress.Parse("A1"));
            Assert.NotNull(merged);
            Assert.Equal(CellRange.Parse("A1:F1"), merged!.Range);
            Assert.Equal(CellAddress.Parse("A1"), merged.Anchor);
        }

        [Fact]
        public void ページ設定と印刷範囲を読み取る()
        {
            var sheet = ReadInvoiceSheet();
            var setup = sheet.PageSetup;

            Assert.Equal("A4", setup.Paper.Name);
            Assert.Equal(PageOrientation.Portrait, setup.Orientation);
            Assert.Equal(PageOrder.DownThenOver, setup.PageOrder);
            Assert.Equal(100, setup.Scaling.ScalePercent);
            Assert.False(setup.Scaling.IsFitToPage);

            // 余白 0.7in = 50.4pt
            Assert.Equal(Units.InchesToPoints(0.7), setup.Margins.LeftPt, precision: 6);

            var printArea = Assert.Single(setup.PrintAreas);
            Assert.Equal(CellRange.Parse("A1:F34"), printArea);
        }

        [Fact]
        public void 印刷タイトルを読み取る()
        {
            var workbook = _reader.ReadFile(TestPaths.SampleTemplate("delivery-note"));
            var sheet = Assert.Single(workbook.Sheets);

            Assert.True(sheet.PageSetup.PrintTitles.HasRows);
            Assert.Equal(1, sheet.PageSetup.PrintTitles.FirstRow);
            Assert.Equal(7, sheet.PageSetup.PrintTitles.LastRow);
            Assert.False(sheet.PageSetup.PrintTitles.HasColumns);
        }

        [Fact]
        public void 複数の印刷範囲を定義順に読み取る()
        {
            // 領収書サンプルは「本紙」と「控え」を別々の印刷範囲として持つ(要件3.6)。
            var workbook = _reader.ReadFile(TestPaths.SampleTemplate("receipt"));
            var sheet = Assert.Single(workbook.Sheets);

            Assert.Equal(2, sheet.PageSetup.PrintAreas.Count);
            Assert.Equal(CellRange.Parse("A1:D13"), sheet.PageSetup.PrintAreas[0]);
            Assert.Equal(CellRange.Parse("A16:D28"), sheet.PageSetup.PrintAreas[1]);
        }

        [Fact]
        public void ヘッダーとフッターの指定を読み取る()
        {
            var workbook = _reader.ReadFile(TestPaths.SampleTemplate("delivery-note"));
            var sheet = Assert.Single(workbook.Sheets);
            var headerFooter = sheet.PageSetup.HeaderFooter;

            Assert.False(headerFooter.IsEmpty);
            Assert.Equal("&R&A", headerFooter.OddHeader);
            Assert.Contains("&P / &N", headerFooter.OddFooter);

            // scaleWithDoc の既定は true
            Assert.True(headerFooter.ScaleWithDocument);
            Assert.False(headerFooter.DifferentFirst);
            Assert.False(headerFooter.DifferentOddEven);
        }

        [Fact]
        public void ヘッダーフッターが無いシートは空として扱う()
        {
            var workbook = _reader.ReadFile(TestPaths.SampleTemplate("invoice"));
            var sheet = Assert.Single(workbook.Sheets);

            Assert.True(sheet.PageSetup.HeaderFooter.IsEmpty);
        }

        [Fact]
        public void シート名で読み取り対象を絞り込める()
        {
            var workbook = _reader.ReadFile(
                TestPaths.SampleTemplate("invoice"),
                new WorkbookReadOptions(SheetNameFilter: "請求書"));

            Assert.Equal("請求書", Assert.Single(workbook.Sheets).Name);
        }

        [Fact]
        public void 存在しないシートを指定したらエラーになる()
        {
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => _reader.ReadFile(
                    TestPaths.SampleTemplate("invoice"),
                    new WorkbookReadOptions(SheetNameFilter: "存在しないシート", ReportCode: "invoice")));

            Assert.Equal(InvalidExcelFileReason.NoWorksheet, ex.Reason);
            Assert.Equal("invoice", ex.ReportCode);
        }

        [Fact]
        public void xlsx以外のファイルは明確なエラーになる()
        {
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes("これはExcelファイルではありません"));

            var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.Read(stream));

            Assert.Equal(InvalidExcelFileReason.NotOpenXmlFormat, ex.Reason);
            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
        }

        [Fact]
        public void パスワード保護されたブックは明確なエラーになる()
        {
            // 暗号化された OOXML は OLE 複合ドキュメント(CFB)でラップされる。
            // そのシグネチャだけを持つストリームで分類を確認する。
            var cfbSignature = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00 };
            using var stream = new MemoryStream(cfbSignature);

            var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.Read(stream));

            Assert.Equal(InvalidExcelFileReason.PasswordProtected, ex.Reason);
        }

        [Fact]
        public void 壊れたZIPは破損として報告される()
        {
            // ZIP のシグネチャだけを持つ壊れたデータ
            var brokenZip = new byte[] { 0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00 };
            using var stream = new MemoryStream(brokenZip);

            var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.Read(stream));

            Assert.True(
                ex.Reason is InvalidExcelFileReason.Corrupted or InvalidExcelFileReason.Unknown,
                $"破損として分類されるべき (実際: {ex.Reason})");
        }

        [Fact]
        public void 存在しないファイルはエラーになる()
        {
            var ex = Assert.Throws<InvalidExcelFileException>(
                () => _reader.ReadFile(Path.Combine(TestPaths.RepositoryRoot, "存在しない.xlsx")));

            Assert.Equal(ProcessingStage.Parsing, ex.Stage);
        }

        [Fact]
        public void シーク不可のストリームからも読み取れる()
        {
            using var file = File.OpenRead(TestPaths.SampleTemplate("invoice"));
            using var nonSeekable = new NonSeekableStream(file);

            var workbook = _reader.Read(nonSeekable);

            Assert.Equal("請求書", Assert.Single(workbook.Sheets).Name);
        }

        /// <summary>シーク不可のストリームを模すラッパー。</summary>
        private sealed class NonSeekableStream : Stream
        {
            private readonly Stream _inner;

            public NonSeekableStream(Stream inner) => _inner = inner;

            public override bool CanRead => true;

            public override bool CanSeek => false;

            public override bool CanWrite => false;

            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush() => _inner.Flush();

            public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
