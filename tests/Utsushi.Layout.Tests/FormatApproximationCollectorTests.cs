using System.Linq;
using Utsushi.Core;
using Xunit;

namespace Utsushi.Layout.Tests
{
    /// <summary>書式の近似の収集器(excel2pdf-consolidation 要件7.4、design.md 3.5)の検証。</summary>
    public sealed class FormatApproximationCollectorTests
    {
        private static readonly CellRange A1 = new(1, 1, 1, 1);

        [Fact]
        public void 同じ種類と位置と要素の識別は1件にまとめる()
        {
            var collector = new FormatApproximationCollector();
            collector.Add(FormatApproximationKind.TextRotation, "S", A1, null, "textRotation=90", "回転");
            collector.Add(FormatApproximationKind.TextRotation, "S", A1, null, "textRotation=90", "回転");
            collector.AddCell(FormatApproximationKind.TextRotation, "S", new CellAddress(1, 1), "textRotation=90", "回転");

            Assert.Single(collector.Items);
        }

        [Fact]
        public void 同じ位置でも別の要素や別の種類は別の件にする()
        {
            var collector = new FormatApproximationCollector();
            collector.Add(FormatApproximationKind.ImageFlip, "S", A1, "id=2", null, "反転");
            collector.Add(FormatApproximationKind.ImageFlip, "S", A1, "id=3", null, "反転");
            collector.Add(FormatApproximationKind.ImageRotation, "S", A1, "id=2", null, "回転");

            Assert.Equal(3, collector.Items.Count);
        }

        [Fact]
        public void 件数は種類ごとに打ち切りほかの種類は記録し続ける()
        {
            var collector = new FormatApproximationCollector(maxPerKind: 2);
            for (var row = 1; row <= 3; row++)
            {
                collector.AddCell(FormatApproximationKind.CellFillPattern, "S", new CellAddress(row, 1), null, "網かけ");
            }

            collector.AddCell(FormatApproximationKind.TextRotation, "S", new CellAddress(1, 2), null, "回転");

            Assert.Equal(2, collector.Items.Count(i => i.Kind == FormatApproximationKind.CellFillPattern));
            Assert.Single(collector.Items, i => i.Kind == FormatApproximationKind.TextRotation);
            Assert.True(collector.IsTruncated);
        }

        [Fact]
        public void 補足の情報は正規化して入れる()
        {
            var collector = new FormatApproximationCollector();
            collector.AddCell(FormatApproximationKind.NumberFormat, "S", new CellAddress(1, 1), "# ?/?\r\n", "表示形式");

            Assert.Equal("# ?/?", collector.Items.Single().Detail);
        }

        [Fact]
        public void 一行の説明にシートと位置を含める()
        {
            Assert.Equal(
                "[ConditionalFormatting] シート 'S' F10:F30: 条件付き書式",
                new FormatApproximation(FormatApproximationKind.ConditionalFormatting, "S", CellRange.Parse("F10:F30"), null, "条件付き書式").ToString());
            Assert.Equal(
                "[TextRotation] シート 'S' D3: 回転",
                new FormatApproximation(FormatApproximationKind.TextRotation, "S", CellRange.Parse("D3:D3"), null, "回転").ToString());
            Assert.Equal(
                "[MultiLineHeaderFooter] シート 'S': 複数行",
                new FormatApproximation(FormatApproximationKind.MultiLineHeaderFooter, "S", null, null, "複数行").ToString());
        }
    }
}
