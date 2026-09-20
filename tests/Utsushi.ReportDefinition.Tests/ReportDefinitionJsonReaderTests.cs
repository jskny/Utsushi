using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.ReportDefinitions;
using Utsushi.ReportDefinitions.Model;
using Xunit;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// 帳票定義JSONのスキーマ検証(要件6.3、タスク3.1 / 8.3)。
    /// </summary>
    public sealed class ReportDefinitionJsonReaderTests
    {
        private const string ValidJson = @"{
      ""schemaVersion"": 1,
      ""reportCode"": ""invoice"",
      ""sheetName"": ""請求書"",
      ""toleranceMm"": 0.3,
      ""unsupportedElements"": ""error"",
      ""maxDigitWidthPx"": 7,
      ""printArea"": ""A1:F34"",
      ""substitutionFields"": [
        { ""key"": ""InvoiceNo"", ""cell"": ""C3"", ""required"": true, ""overflow"": ""shrink"" },
        { ""key"": ""IssueDate"", ""cell"": ""C4"" }
      ]
    }";

        [Fact]
        public void 正しい定義を読み込める()
        {
            var definition = ReportDefinitionJsonReader.Read(ValidJson);

            Assert.Equal("invoice", definition.ReportCode);
            Assert.Equal("請求書", definition.SheetName);
            Assert.Equal(0.3, definition.ToleranceMm);
            Assert.Equal(UnsupportedElementPolicy.Error, definition.UnsupportedElements);
            Assert.Equal(7.0, definition.MaxDigitWidthPx);
            Assert.Equal(CellRange.Parse("A1:F34"), definition.PrintAreaOverride);
            Assert.Equal(2, definition.SubstitutionFields.Count);

            Assert.True(definition.TryGetField("InvoiceNo", out var invoiceNo));
            Assert.Equal(CellAddress.Parse("C3"), invoiceNo.Cell);
            Assert.True(invoiceNo.Required);
            Assert.Equal(OverflowBehavior.Shrink, invoiceNo.Overflow);
        }

        [Fact]
        public void 省略可能な項目には既定値が入る()
        {
            var json = @"{
          ""schemaVersion"": 1,
          ""reportCode"": ""minimal"",
          ""sheetName"": ""Sheet1"",
          ""substitutionFields"": []
        }";

            var definition = ReportDefinitionJsonReader.Read(json);

            Assert.Equal(ReportDefinitions.Model.ReportDefinition.DefaultToleranceMm, definition.ToleranceMm);
            Assert.Equal(ReportDefinitions.Model.ReportDefinition.DefaultMaxDigitWidthPx, definition.MaxDigitWidthPx);
            Assert.Equal(UnsupportedElementPolicy.Ignore, definition.UnsupportedElements);
            Assert.Null(definition.PrintAreaOverride);

            // required を省略したフィールドは任意扱い。
            // overflow を省略した場合は null(= Excel 側のセル書式に従う)になる。
            var d2 = ReportDefinitionJsonReader.Read(ValidJson);
            Assert.True(d2.TryGetField("IssueDate", out var issueDate));
            Assert.False(issueDate.Required);
            Assert.Null(issueDate.Overflow);
        }

        [Fact]
        public void JSONとして壊れていればスキーマ例外になる()
        {
            var ex = Assert.Throws<ReportDefinitionSchemaException>(
                () => ReportDefinitionJsonReader.Read("{ これはJSONではない }", "/path/to/definition.json"));

            Assert.Equal("/path/to/definition.json", ex.DefinitionPath);
            Assert.Equal(ProcessingStage.ReportDefinition, ex.Stage);
        }

        [Theory]
        [InlineData(@"{ ""reportCode"": ""x"", ""sheetName"": ""s"", ""substitutionFields"": [] }", "schemaVersion")]
        [InlineData(@"{ ""schemaVersion"": 1, ""sheetName"": ""s"", ""substitutionFields"": [] }", "reportCode")]
        [InlineData(@"{ ""schemaVersion"": 1, ""reportCode"": ""x"", ""substitutionFields"": [] }", "sheetName")]
        [InlineData(@"{ ""schemaVersion"": 1, ""reportCode"": ""x"", ""sheetName"": ""s"" }", "substitutionFields")]
        public void 必須プロパティが無ければ該当プロパティ名つきで失敗する(string json, string expectedPropertyPath)
        {
            var ex = Assert.Throws<ReportDefinitionSchemaException>(() => ReportDefinitionJsonReader.Read(json));

            Assert.Equal(expectedPropertyPath, ex.PropertyPath);
        }

        [Fact]
        public void 対応していないスキーマバージョンは拒否する()
        {
            var json = @"{ ""schemaVersion"": 99, ""reportCode"": ""x"", ""sheetName"": ""s"", ""substitutionFields"": [] }";

            var ex = Assert.Throws<ReportDefinitionSchemaException>(() => ReportDefinitionJsonReader.Read(json));

            Assert.Equal("schemaVersion", ex.PropertyPath);
            Assert.Contains("99", ex.Message);
        }

        [Fact]
        public void セル番地として解釈できない値は拒否する()
        {
            var json = @"{
          ""schemaVersion"": 1, ""reportCode"": ""x"", ""sheetName"": ""s"",
          ""substitutionFields"": [ { ""key"": ""K"", ""cell"": ""セル3"" } ]
        }";

            var ex = Assert.Throws<ReportDefinitionSchemaException>(() => ReportDefinitionJsonReader.Read(json));

            Assert.Equal("substitutionFields[0].cell", ex.PropertyPath);
        }

        [Fact]
        public void 置換キーの重複は拒否する()
        {
            var json = @"{
          ""schemaVersion"": 1, ""reportCode"": ""x"", ""sheetName"": ""s"",
          ""substitutionFields"": [
            { ""key"": ""K"", ""cell"": ""A1"" },
            { ""key"": ""K"", ""cell"": ""A2"" }
          ]
        }";

            var ex = Assert.Throws<ReportDefinitionSchemaException>(() => ReportDefinitionJsonReader.Read(json));

            Assert.Equal("substitutionFields[1].key", ex.PropertyPath);
        }

        [Fact]
        public void overflowの値が不正なら拒否する()
        {
            var json = @"{
          ""schemaVersion"": 1, ""reportCode"": ""x"", ""sheetName"": ""s"",
          ""substitutionFields"": [ { ""key"": ""K"", ""cell"": ""A1"", ""overflow"": ""squeeze"" } ]
        }";

            var ex = Assert.Throws<ReportDefinitionSchemaException>(() => ReportDefinitionJsonReader.Read(json));

            Assert.Equal("substitutionFields[0].overflow", ex.PropertyPath);
            Assert.Contains("squeeze", ex.Message);
        }

        [Fact]
        public void unsupportedElementsの値が不正なら拒否する()
        {
            var json = @"{
          ""schemaVersion"": 1, ""reportCode"": ""x"", ""sheetName"": ""s"",
          ""unsupportedElements"": ""warn"", ""substitutionFields"": []
        }";

            var ex = Assert.Throws<ReportDefinitionSchemaException>(() => ReportDefinitionJsonReader.Read(json));

            Assert.Equal("unsupportedElements", ex.PropertyPath);
        }

        [Theory]
        [InlineData(@"""toleranceMm"": -1", "toleranceMm")]
        [InlineData(@"""maxDigitWidthPx"": 0", "maxDigitWidthPx")]
        [InlineData(@"""toleranceMm"": ""0.5""", "toleranceMm")]
        [InlineData(@"""printArea"": ""not-a-range""", "printArea")]
        public void 値域や型が不正なら拒否する(string property, string expectedPropertyPath)
        {
            var json = $@"{{
          ""schemaVersion"": 1, ""reportCode"": ""x"", ""sheetName"": ""s"",
          {property}, ""substitutionFields"": []
        }}";

            var ex = Assert.Throws<ReportDefinitionSchemaException>(() => ReportDefinitionJsonReader.Read(json));

            Assert.Equal(expectedPropertyPath, ex.PropertyPath);
        }

        [Fact]
        public void リポジトリの帳票定義が全て読み込める()
        {
            // samples/reports 配下の実際の帳票定義がスキーマを満たすことを確認する(要件8.3)。
            var root = TestPaths.SampleReportsRoot;
            var repository = new FileSystemReportDefinitionRepository(root);

            var codes = repository.ListReportCodes();
            Assert.NotEmpty(codes);

            foreach (var code in codes)
            {
                var definition = repository.Load(code);
                Assert.Equal(code, definition.ReportCode);
                Assert.NotEmpty(definition.SheetName);
            }
        }
    }
}
