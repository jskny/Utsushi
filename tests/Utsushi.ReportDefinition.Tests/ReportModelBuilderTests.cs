using System;
using System.Collections.Generic;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Utsushi.ReportDefinitions;
using Utsushi.ReportDefinitions.Model;
using Xunit;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// 変換の設定とワークブックからの帳票モデルの構築の検証。
    /// </summary>
    public sealed class ReportModelBuilderTests
    {
        private readonly ReportModelBuilder _builder = new();

        private static ReportDefinitions.Model.ConversionSettings Definition(string sheetName) =>
            new("test", sheetName, ReportDefinitions.Model.ConversionSettings.DefaultMaxDigitWidthPx);

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
        public void 設定のシートを取り出してモデルを構築する()
        {
            var definition = Definition("請求書");
            var workbook = Workbook("請求書", new[] { "A1", "C3", "F10" });

            var report = _builder.Build(workbook, definition);

            Assert.Same(definition, report.Settings);
            Assert.Equal("請求書", report.Sheet.Name);
        }

        [Fact]
        public void 設定のシートが無ければ実在するシート名を含むInvalidOperationExceptionにする()
        {
            // 読み込みは対象のシートだけを返すため通常は起きない。契約に反する読み込みの実装が差し替えられた場合の防御。
            var workbook = Workbook("見積書", new[] { "A1" });

            var ex = Assert.Throws<InvalidOperationException>(() => _builder.Build(workbook, Definition("請求書")));

            Assert.Contains("請求書", ex.Message);
            Assert.Contains("見積書", ex.Message);
        }
    }
}
