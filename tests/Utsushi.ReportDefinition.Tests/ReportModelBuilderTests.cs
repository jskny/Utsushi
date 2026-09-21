using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions;
using Utsushi.ReportDefinitions.Model;
using Xunit;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// 帳票定義とワークブックの突合(要件1.4、タスク3.2)の検証。
    /// </summary>
    public sealed class ReportModelBuilderTests
    {
        private readonly ReportModelBuilder _builder = new();

        private static ReportDefinitions.Model.ReportDefinition Definition(
            string sheetName, params (string Key, string Cell)[] fields)
        {
            var list = new List<SubstitutionFieldDefinition>();
            foreach (var (key, cell) in fields)
            {
                list.Add(new SubstitutionFieldDefinition(key, CellAddress.Parse(cell), false, Overflow: null));
            }

            return new ReportDefinitions.Model.ReportDefinition(
                "test", sheetName, list,
                ReportDefinitions.Model.ReportDefinition.DefaultToleranceMm,
                UnsupportedElementPolicy.Ignore,
                ReportDefinitions.Model.ReportDefinition.DefaultMaxDigitWidthPx,
                null);
        }

        private static WorkbookModel Workbook(
            string sheetName, IEnumerable<string> cellAddresses, params CellRange[] mergedRanges)
        {
            var cells = new Dictionary<CellAddress, CellModel>();
            foreach (var address in cellAddresses)
            {
                cells[CellAddress.Parse(address)] = new CellModel("値", CellValueKind.Text, CellStyle.Default, "値");
            }

            var merged = new List<MergedRange>();
            foreach (var range in mergedRanges)
            {
                merged.Add(new MergedRange(range));
            }

            var sheet = new SheetModel(
                sheetName, cells, merged,
                new List<double>(), new List<double>(),
                8.43, 15.0,
                new HashSet<int>(), new HashSet<int>(),
                PageSetupModel.Default,
                Array.Empty<DrawingObjectModel>());

            return new WorkbookModel(new[] { sheet }, FontStyle.Default);
        }

        [Fact]
        public void 定義とシートが一致すればモデルを構築できる()
        {
            var definition = Definition("請求書", ("InvoiceNo", "C3"));
            var workbook = Workbook("請求書", new[] { "A1", "C3", "F10" });

            var report = _builder.Build(workbook, definition);

            Assert.Same(definition, report.Definition);
            Assert.Equal("請求書", report.Sheet.Name);
            Assert.Empty(report.OverflowByCell);
        }

        [Fact]
        public void 期待するシートが無ければエラーになる()
        {
            var definition = Definition("請求書");
            var workbook = Workbook("見積書", new[] { "A1" });

            var ex = Assert.Throws<ReportStructureMismatchException>(() => _builder.Build(workbook, definition));

            Assert.Equal("請求書", ex.SheetName);
            Assert.Equal("test", ex.ReportCode);
            Assert.Equal(ProcessingStage.ReportDefinition, ex.Stage);

            // 実在するシート名をメッセージに含め、原因を特定できるようにする(要件6.4)。
            Assert.Contains("見積書", ex.Message);
        }

        [Fact]
        public void 置換対象セルが使用範囲の外ならエラーになる()
        {
            var definition = Definition("請求書", ("Far", "Z100"));
            var workbook = Workbook("請求書", new[] { "A1", "C3" });

            var ex = Assert.Throws<ReportStructureMismatchException>(() => _builder.Build(workbook, definition));

            Assert.Equal(CellAddress.Parse("Z100"), ex.CellAddress);
            Assert.Contains("Far", ex.Message);
        }

        [Fact]
        public void 置換対象セルが結合範囲の左上なら受け付ける()
        {
            var definition = Definition("請求書", ("Title", "A1"));
            var workbook = Workbook("請求書", new[] { "A1", "F10" }, CellRange.Parse("A1:F1"));

            var report = _builder.Build(workbook, definition);

            Assert.NotNull(report);
        }

        [Fact]
        public void 置換対象セルが結合範囲の内側ならエラーになる()
        {
            // B1 は A1:F1 の内側。Excel 上で値を持てるのは左上の A1 だけなので拒否する。
            var definition = Definition("請求書", ("Title", "B1"));
            var workbook = Workbook("請求書", new[] { "A1", "B1", "F10" }, CellRange.Parse("A1:F1"));

            var ex = Assert.Throws<ReportStructureMismatchException>(() => _builder.Build(workbook, definition));

            Assert.Equal(CellAddress.Parse("B1"), ex.CellAddress);
            Assert.Contains("A1", ex.Message);
        }

        [Fact]
        public void 値が空のセルでも使用範囲内なら受け付ける()
        {
            // 置換先が空セル(Cells に現れない)でも、使用範囲に含まれていれば有効とする。
            var definition = Definition("請求書", ("Empty", "C3"));
            var workbook = Workbook("請求書", new[] { "A1", "F10" });

            var report = _builder.Build(workbook, definition);

            Assert.Null(report.Sheet.GetCell(CellAddress.Parse("C3")));
        }
    }
}
