using Utsushi.Core;
using Utsushi.ReportDefinitions.Model;
using Xunit;
using Definition = Utsushi.ReportDefinitions.Model.ReportDefinition;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// <see cref="Definition.TryGetField"/> の索引(遅延キャッシュ)が、record の <c>with</c> と等値比較に
    /// 巻き込まれないことを確かめる。
    /// </summary>
    public sealed class FieldIndexCacheTests
    {
        private static readonly SubstitutionFieldDefinition[] OriginalFields =
        {
            new("InvoiceNo", CellAddress.Parse("C3"), Required: true, Overflow: null),
        };

        private static Definition Create() =>
            new(
                "invoice",
                "請求書",
                OriginalFields,
                Definition.DefaultToleranceMm,
                UnsupportedElementPolicy.Ignore,
                Definition.DefaultMaxDigitWidthPx,
                PrintAreaOverride: null);

        [Fact]
        public void withで置換フィールドを差し替えた後は新しいフィールドで引ける()
        {
            var original = Create();

            // 索引を構築させてから with する。
            Assert.True(original.TryGetField("InvoiceNo", out _));

            var replaced = original with
            {
                SubstitutionFields = new[] { new SubstitutionFieldDefinition("CustomerName", CellAddress.Parse("B2"), false, null) },
            };

            Assert.True(replaced.TryGetField("CustomerName", out var field));
            Assert.Equal(CellAddress.Parse("B2"), field.Cell);
            Assert.False(replaced.TryGetField("InvoiceNo", out _));

            // 元の定義の索引は変わらない。
            Assert.True(original.TryGetField("InvoiceNo", out _));
            Assert.False(original.TryGetField("CustomerName", out _));
        }

        [Fact]
        public void 索引を構築済みかどうかは等値比較に影響しない()
        {
            var built = Create();
            var notBuilt = Create();
            Assert.True(built.TryGetField("InvoiceNo", out _));

            Assert.Equal(notBuilt, built);
            Assert.True(built == notBuilt);
            Assert.Equal(notBuilt.GetHashCode(), built.GetHashCode());
        }

        [Fact]
        public void withで他の項目を変えた定義とは等しくない()
        {
            var original = Create();

            Assert.NotEqual(original, original with { SheetName = "別シート" });
            Assert.NotEqual(original, original with { PrintAreaOverride = CellRange.Parse("A1:B2") });
            Assert.Equal(original, original with { });
        }
    }
}
