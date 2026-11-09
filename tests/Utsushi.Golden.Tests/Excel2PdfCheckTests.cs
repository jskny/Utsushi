using System;
using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Rendering;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// 診断 <see cref="Excel2Pdf.Check"/>(excel2pdf-consolidation 要件3・4)。
    /// </summary>
    public sealed class Excel2PdfCheckTests
    {
        private static readonly string Invoice = TestPaths.SampleTemplate("invoice");
        private static readonly string OrderForm = Path.Combine(TestPaths.EdgeCaseSamplesRoot, "edge-order-form", "template.xlsx");

        private static Excel2PdfOptions FallbackFonts() => new() { Fonts = FontResolverOptions.AllowFallback() };

        /// <summary>実行環境にフォントがインストールされているかどうか(フォントの代替が起きるかは環境によって変わるため)。</summary>
        private static bool IsInstalled(string family)
        {
            using var resolver = new FontResolver(FontResolverOptions.Strict);
            try
            {
                resolver.Resolve(FontStyle.Default with { Name = family });
                return true;
            }
            catch (FontNotAvailableException)
            {
                return false;
            }
        }

        [Fact]
        public void 未対応要素の扱いがエラーでも例外にせず一覧で返す()
        {
            using var pdf = new Excel2Pdf(OrderForm, FallbackFonts());

            var result = pdf.Check();

            Assert.Contains(result.UnsupportedElements, e => e.Kind == UnsupportedElementKind.Chart);
            Assert.Contains(result.UnsupportedElements, e => e.Kind == UnsupportedElementKind.LegacyDrawing);
            Assert.False(result.UnsupportedElementsTruncated);
            Assert.True(result.HasIssues);
        }

        [Fact]
        public void 文字の収まらない箇所を設定した値で返す()
        {
            using var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.SetValue("F12", 123456789012345);

            var result = pdf.Check();

            Assert.Contains(
                result.FitIssues,
                i => i.Kind == Layout.Model.FitIssueKind.NumberTooWide && i.Cell == CellAddress.Parse("F12") && i.IsSubstituted);
            Assert.True(result.HasIssues);
        }

        [Fact]
        public void 未対応要素以外のエラーは保存と同じエラーにする()
        {
            using var pdf = new Excel2Pdf(OrderForm, FallbackFonts());
            pdf.SetText("A3", "\u0001");

            Assert.Throws<InvalidSubstitutionValueException>(() => pdf.Check());

            // 未対応要素の扱いを無視にした保存と同じエラー(要件4.4)。
            using var ignoring = new Excel2Pdf(
                OrderForm,
                new Excel2PdfOptions { Fonts = FontResolverOptions.AllowFallback(), UnsupportedElements = UnsupportedElementPolicy.Ignore });
            ignoring.SetText("A3", "\u0001");
            Assert.Throws<InvalidSubstitutionValueException>(() => ignoring.ToPdfBytes());
        }

        [Fact]
        public void 代替したフォントを返し同じエンジンの2回目の変換でも返す()
        {
            using var engine = Excel2PdfEngine.Create(FontResolverOptions.AllowFallback());
            using var first = new Excel2Pdf(Invoice, engine);
            using var second = new Excel2Pdf(Invoice, engine);

            var firstResult = first.Check().FontSubstitutions;
            var secondResult = second.Check().FontSubstitutions;

            Assert.Equal(firstResult, secondResult);
            if (!IsInstalled("MS PGothic"))
            {
                var substitution = Assert.Single(firstResult, s => s.RequestedFamily == "MS PGothic" && !s.Bold && !s.Italic);
                Assert.Equal(Rendering.Fonts.BundledFonts.JapaneseGothicFamily, substitution.ActualFamily);
                Assert.Equal("MS PGothic → " + Rendering.Fonts.BundledFonts.JapaneseGothicFamily, substitution.ToString());
            }
        }

        [Fact]
        public void その変換で使っていないフォントは返さない()
        {
            using var engine = Excel2PdfEngine.Create(FontResolverOptions.AllowFallback());
            using (var orderForm = new Excel2Pdf(OrderForm, engine, new Excel2PdfOptions { UnsupportedElements = UnsupportedElementPolicy.Ignore }))
            {
                orderForm.Check();
            }

            using var invoice = new Excel2Pdf(Invoice, engine);
            var result = invoice.Check();

            Assert.DoesNotContain(result.FontSubstitutions, s => s.RequestedFamily.Contains("游", StringComparison.Ordinal));
            Assert.Equal(result.FontSubstitutions.Count, result.FontSubstitutions.Distinct().Count());
        }

        [Fact]
        public void 問題が無ければ問題無しと判定できる()
        {
            var empty = new Excel2PdfDiagnostics(
                Array.Empty<Layout.Model.FitIssue>(), false, Array.Empty<UnsupportedElement>(), false, Array.Empty<FontSubstitution>(),
                Array.Empty<FormatApproximation>(), false);

            Assert.False(empty.HasIssues);
            Assert.False(empty.HasFormatApproximations);
            var approximated = empty with
            {
                FormatApproximations = new[] { new FormatApproximation(FormatApproximationKind.TextRotation, "S", null, null, "回転") },
            };
            Assert.False(approximated.HasIssues);
            Assert.True(approximated.HasFormatApproximations);
            Assert.True((empty with { FontSubstitutions = new[] { new FontSubstitution("A", true, false, "B") } }).HasIssues);
            Assert.Equal("A(太字) → B", new FontSubstitution("A", true, false, "B").ToString());
        }

        [Fact]
        public void 破棄した後は使えない()
        {
            var pdf = new Excel2Pdf(Invoice, FallbackFonts());
            pdf.Dispose();

            Assert.Throws<ObjectDisposedException>(() => pdf.Check());
        }
    }
}
