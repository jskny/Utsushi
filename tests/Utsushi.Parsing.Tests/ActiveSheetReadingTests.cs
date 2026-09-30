using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Spec = Utsushi.Parsing.Tests.ActiveSheetWorkbookFixtures.SheetSpec;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 帳票定義なしモードのアクティブシート解決(<see cref="WorkbookReadOptions.ActiveSheetOnly"/>)の検証
    /// (要件12.2、タスク24)。
    /// </summary>
    public sealed class ActiveSheetReadingTests
    {
        private static readonly WorkbookReadOptions ActiveOnly = new(ActiveSheetOnly: true);

        private readonly OpenXmlWorkbookReader _reader = new();

        /// <summary>ブックを作成して読み取り、選ばれたシート名の一覧を返す。</summary>
        private string[] ReadSheetNames(WorkbookReadOptions options, uint? activeTab, params Spec[] sheets)
        {
            var path = ActiveSheetWorkbookFixtures.CreateWorkbook(activeTab, sheets);
            try
            {
                using var stream = File.OpenRead(path);
                return _reader.Read(stream, options).Sheets.Select(s => s.Name).ToArray();
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ActiveSheetOnlyを指定しなければ全シートを読む()
        {
            // 前提の確認: 既定では従来どおり全シートを読む(ActiveSheetOnly の既定値は false)。
            Assert.False(WorkbookReadOptions.Default.ActiveSheetOnly);

            var names = ReadSheetNames(
                WorkbookReadOptions.Default, 1U, new Spec("表紙"), new Spec("明細"), new Spec("集計"));

            Assert.Equal(new[] { "表紙", "明細", "集計" }, names);
        }

        [Fact]
        public void activeTabがグラフシートを指すと表示されている最初のワークシートを読む()
        {
            // グラフシートはワークシートとして読めないため候補にしない(code-reviewer等の指摘)。
            // 以前は選んだグラフシートが読み取りループで読み飛ばされ、誤って NoWorksheet になっていた。
            var names = ReadSheetNames(
                ActiveOnly, 0U, new Spec("グラフ1", IsChartSheet: true), new Spec("明細"), new Spec("集計"));

            Assert.Equal(new[] { "明細" }, names);
        }

        [Fact]
        public void 表示されている最初のシートがグラフシートなら次のワークシートを読む()
        {
            var names = ReadSheetNames(
                ActiveOnly,
                0U,
                new Spec("表紙", SheetStateValues.Hidden),
                new Spec("グラフ1", IsChartSheet: true),
                new Spec("明細"));

            Assert.Equal(new[] { "明細" }, names);
        }

        [Fact]
        public void グラフシートしか表示されていなければNoWorksheetエラーになる()
        {
            var path = ActiveSheetWorkbookFixtures.CreateWorkbook(
                0U, new Spec("グラフ1", IsChartSheet: true), new Spec("明細", SheetStateValues.Hidden));
            try
            {
                using var stream = File.OpenRead(path);
                var ex = Assert.Throws<InvalidExcelFileException>(() => _reader.Read(stream, ActiveOnly));
                Assert.Equal(InvalidExcelFileReason.NoWorksheet, ex.Reason);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void activeTabが2番目のシートを指すとそのシートだけを読む()
        {
            var path = ActiveSheetWorkbookFixtures.CreateWorkbook(
                1U, new Spec("表紙"), new Spec("明細"), new Spec("集計"));
            try
            {
                using var stream = File.OpenRead(path);
                var sheet = Assert.Single(_reader.Read(stream, ActiveOnly).Sheets);

                Assert.Equal("明細", sheet.Name);
                // シート名だけでなく、中身もそのシートのものであること。
                Assert.Equal("明細", sheet.GetCell(CellAddress.Parse("A1"))!.DisplayValue);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void bookViewsが無ければ先頭のシートを読む()
        {
            var names = ReadSheetNames(ActiveOnly, activeTab: null, new Spec("表紙"), new Spec("明細"));

            Assert.Equal(new[] { "表紙" }, names);
        }

        [Fact]
        public void activeTabが0なら先頭のシートを読む()
        {
            var names = ReadSheetNames(ActiveOnly, 0U, new Spec("表紙"), new Spec("明細"));

            Assert.Equal(new[] { "表紙" }, names);
        }

        [Theory]
        [InlineData("hidden")]
        [InlineData("veryHidden")]
        public void activeTabが非表示シートを指すと表示されている最初のシートを読む(string state)
        {
            var names = ReadSheetNames(
                ActiveOnly,
                2U,
                new Spec("非表示1", ParseState(state)),
                new Spec("表示1"),
                new Spec("非表示2", ParseState(state)),
                new Spec("表示2"));

            Assert.Equal(new[] { "表示1" }, names);
        }

        [Theory]
        [InlineData("hidden")]
        [InlineData("veryHidden")]
        public void activeTab省略時に先頭が非表示なら表示されている最初のシートを読む(string state)
        {
            var names = ReadSheetNames(
                ActiveOnly,
                activeTab: null,
                new Spec("非表示", ParseState(state)),
                new Spec("表示1"),
                new Spec("表示2"));

            Assert.Equal(new[] { "表示1" }, names);
        }

        [Fact]
        public void 明示的にvisibleのシートは表示シートとして扱う()
        {
            var names = ReadSheetNames(
                ActiveOnly, 1U, new Spec("表紙"), new Spec("明細", SheetStateValues.Visible));

            Assert.Equal(new[] { "明細" }, names);
        }

        [Fact]
        public void activeTabがシート数以上なら表示されている最初のシートを読む()
        {
            var names = ReadSheetNames(
                ActiveOnly, 5U, new Spec("非表示", SheetStateValues.Hidden), new Spec("表紙"), new Spec("明細"));

            Assert.Equal(new[] { "表紙" }, names);
        }

        [Fact]
        public void 表示されているシートが無ければNoWorksheetエラーになる()
        {
            var path = ActiveSheetWorkbookFixtures.CreateWorkbook(
                0U,
                new Spec("非表示", SheetStateValues.Hidden),
                new Spec("完全非表示", SheetStateValues.VeryHidden));
            try
            {
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<InvalidExcelFileException>(
                    () => _reader.Read(stream, new WorkbookReadOptions(ReportCode: "文書", ActiveSheetOnly: true)));

                Assert.Equal(InvalidExcelFileReason.NoWorksheet, ex.Reason);
                Assert.Equal(ProcessingStage.Parsing, ex.Stage);
                Assert.Equal("文書", ex.ReportCode);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void SheetNameFilterとActiveSheetOnlyの両方を指定するとSheetNameFilterが優先される()
        {
            var names = ReadSheetNames(
                new WorkbookReadOptions(SheetNameFilter: "集計", ActiveSheetOnly: true),
                1U,
                new Spec("表紙"),
                new Spec("明細"),
                new Spec("集計"));

            Assert.Equal(new[] { "集計" }, names);
        }

        [Fact]
        public void SheetNameFilterが優先されるので非表示シートも名前指定なら読める()
        {
            var names = ReadSheetNames(
                new WorkbookReadOptions(SheetNameFilter: "非表示", ActiveSheetOnly: true),
                0U,
                new Spec("表紙"),
                new Spec("非表示", SheetStateValues.Hidden));

            Assert.Equal(new[] { "非表示" }, names);
        }

        [Fact]
        public void 選ばれなかったシートに非対応の図形があってもunsupportedElementsがerrorで失敗しない()
        {
            var unsupported = new OpenXmlElement[]
            {
                ShapeWorkbookFixtures.ShapeAnchor(A.ShapeTypeValues.FlowChartPreparation),
            };

            var path = ActiveSheetWorkbookFixtures.CreateWorkbook(
                0U, new Spec("帳票"), new Spec("作業用", Anchors: unsupported));
            try
            {
                // 前提の確認: 非対応の図形があるシートを読めば error で失敗する入力であること。
                using (var stream = File.OpenRead(path))
                {
                    var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(
                        stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error, SheetNameFilter: "作業用")));
                    Assert.Equal("UnsupportedShapePreset", ex.ElementKind);
                }

                using (var stream = File.OpenRead(path))
                {
                    var workbook = _reader.Read(
                        stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error, ActiveSheetOnly: true));
                    Assert.Equal("帳票", Assert.Single(workbook.Sheets).Name);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 選ばれなかったシートの図形数が上限を超えていてもunsupportedElementsがerrorで失敗しない()
        {
            var tooMany = Enumerable.Range(0, OpenXmlWorkbookReader.MaxShapesPerSheet + 1)
                .Select(i => (OpenXmlElement)ShapeWorkbookFixtures.ShapeAnchor(
                    A.ShapeTypeValues.Rectangle, row: 1 + i, id: (uint)(2 + i)))
                .ToArray();

            var path = ActiveSheetWorkbookFixtures.CreateWorkbook(
                1U, new Spec("作業用", Anchors: tooMany), new Spec("帳票"));
            try
            {
                // 前提の確認: 図形数が上限を超えたシートを読めば error で失敗する入力であること。
                using (var stream = File.OpenRead(path))
                {
                    var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(
                        stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error, SheetNameFilter: "作業用")));
                    Assert.Equal("TooManyShapes", ex.ElementKind);
                }

                using (var stream = File.OpenRead(path))
                {
                    var workbook = _reader.Read(
                        stream, new WorkbookReadOptions(UnsupportedElementBehavior.Error, ActiveSheetOnly: true));
                    Assert.Equal("帳票", Assert.Single(workbook.Sheets).Name);
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        private static SheetStateValues ParseState(string state) => state switch
        {
            "hidden" => SheetStateValues.Hidden,
            "veryHidden" => SheetStateValues.VeryHidden,
            _ => throw new ArgumentOutOfRangeException(nameof(state), state, null),
        };
    }
}
