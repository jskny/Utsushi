using System;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using Definition = Utsushi.ReportDefinitions.Model.ReportDefinition;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// 帳票定義なしモードで使う既定値の帳票定義(<see cref="Definition.CreateWithoutDefinition"/>)の検証
    /// (要件12、タスク24)。
    /// </summary>
    public sealed class CreateWithoutDefinitionTests
    {
        [Fact]
        public void 文書名とシート名以外は既定値になる()
        {
            var definition = Definition.CreateWithoutDefinition("見積書_2026", "見積");

            Assert.Equal("見積書_2026", definition.ReportCode);
            Assert.Equal("見積", definition.SheetName);
            Assert.Empty(definition.SubstitutionFields);
            Assert.Equal(Definition.DefaultToleranceMm, definition.ToleranceMm);
            Assert.Equal(UnsupportedElementPolicy.Ignore, definition.UnsupportedElements);
            Assert.Equal(Definition.DefaultMaxDigitWidthPx, definition.MaxDigitWidthPx);
            Assert.Null(definition.PrintAreaOverride);
        }

        [Fact]
        public void 文書名は空文字列でもよい()
        {
            // ファサードは documentName が null の場合に空文字列を渡す。
            var definition = Definition.CreateWithoutDefinition(string.Empty, "Sheet1");

            Assert.Equal(string.Empty, definition.ReportCode);
            Assert.Equal("Sheet1", definition.SheetName);
        }

        [Fact]
        public void 置換キーを持たないのでどのキーも引けない()
        {
            var definition = Definition.CreateWithoutDefinition("doc", "Sheet1");

            Assert.False(definition.TryGetField("CustomerName", out _));
        }

        [Fact]
        public void 文書名がnullならArgumentNullException()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => Definition.CreateWithoutDefinition(null!, "Sheet1"));

            Assert.Equal("documentName", ex.ParamName);
        }

        [Fact]
        public void シート名がnullならArgumentNullException()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => Definition.CreateWithoutDefinition("doc", null!));

            Assert.Equal("sheetName", ex.ParamName);
        }
    }
}
