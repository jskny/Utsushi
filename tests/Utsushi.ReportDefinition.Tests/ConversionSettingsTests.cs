using System;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using Definition = Utsushi.ReportDefinitions.Model.ConversionSettings;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// 変換の設定(<see cref="Definition.Create"/>)の検証(要件12)。
    /// </summary>
    public sealed class ConversionSettingsTests
    {
        [Fact]
        public void 最大数字幅を省略すると既定値になる()
        {
            var definition = Definition.Create("見積書_2026", "見積");

            Assert.Equal("見積書_2026", definition.DocumentName);
            Assert.Equal("見積", definition.SheetName);
            Assert.Equal(Definition.DefaultMaxDigitWidthPx, definition.MaxDigitWidthPx);
        }

        [Fact]
        public void 文書名は空文字列でもよい()
        {
            // ファサードは documentName が null の場合に空文字列を渡す。
            var definition = Definition.Create(string.Empty, "Sheet1");

            Assert.Equal(string.Empty, definition.DocumentName);
            Assert.Equal("Sheet1", definition.SheetName);
        }

        [Fact]
        public void 文書名がnullならArgumentNullException()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => Definition.Create(null!, "Sheet1"));

            Assert.Equal("documentName", ex.ParamName);
        }

        [Fact]
        public void シート名がnullならArgumentNullException()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => Definition.Create("doc", null!));

            Assert.Equal("sheetName", ex.ParamName);
        }
    }
}
