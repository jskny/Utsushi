using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 未対応要素の収集器(excel2pdf-consolidation 要件2.4・2.5)と、診断の文字列の正規化(design.md 2.1)の検証。
    /// </summary>
    public sealed class UnsupportedElementCollectorTests
    {
        [Fact]
        public void 要素ごとに記録し記録した順に返す()
        {
            var collector = new UnsupportedElementCollector();
            collector.Add(UnsupportedElementKind.Chart, "注文書", CellAddress.Parse("D10"), null, "グラフは出力されません");
            collector.Add(UnsupportedElementKind.Chart, "注文書", CellAddress.Parse("D30"), null, "グラフは出力されません");

            Assert.Equal(2, collector.Elements.Count);
            Assert.Equal(CellAddress.Parse("D10"), collector.Elements[0].Anchor);
            Assert.Equal(CellAddress.Parse("D30"), collector.Elements[1].Anchor);
            Assert.False(collector.IsTruncated);
        }

        [Fact]
        public void 枚数の上限の超過はシートと上限の種類ごとに1件だけ記録する()
        {
            var collector = new UnsupportedElementCollector();
            collector.AddCountLimitExceeded("Sheet1", CellAddress.Parse("A1"), "TooManyShapes", 2000, "図形の数が上限を超えました");
            collector.AddCountLimitExceeded("Sheet1", CellAddress.Parse("B2"), "TooManyShapes", 2000, "図形の数が上限を超えました");
            collector.AddCountLimitExceeded("Sheet1", null, "TooManyImages", 50, "画像の枚数が上限を超えました");
            collector.AddCountLimitExceeded("Sheet2", null, "TooManyShapes", 2000, "図形の数が上限を超えました");

            Assert.Equal(3, collector.Elements.Count);
            var first = collector.Elements[0];
            Assert.Equal(UnsupportedElementKind.LimitExceeded, first.Kind);
            Assert.Equal("TooManyShapes: 上限 2000", first.Detail);
            Assert.Equal(CellAddress.Parse("A1"), first.Anchor);
        }

        [Fact]
        public void 一つの要素の上限の超過は要素ごとに記録する()
        {
            var collector = new UnsupportedElementCollector();
            collector.AddElementLimitExceeded("Sheet1", CellAddress.Parse("A1"), "ImageTooLarge", 20, "画像が大きすぎます");
            collector.AddElementLimitExceeded("Sheet1", CellAddress.Parse("A5"), "ImageTooLarge", 20, "画像が大きすぎます");

            Assert.Equal(2, collector.Elements.Count);
            Assert.All(collector.Elements, e => Assert.Equal("ImageTooLarge: 上限 20", e.Detail));
        }

        [Fact]
        public void 件数の上限に達したら記録せず打ち切りの印を立てる()
        {
            var collector = new UnsupportedElementCollector(maxElements: 2);
            for (var i = 1; i <= 3; i++)
            {
                collector.Add(UnsupportedElementKind.LegacyDrawing, "Sheet" + i, null, null, "コメント");
            }

            Assert.Equal(2, collector.Elements.Count);
            Assert.True(collector.IsTruncated);
        }

        [Fact]
        public void 件数の上限ちょうどでは打ち切りの印を立てない()
        {
            var collector = new UnsupportedElementCollector(maxElements: 2);
            collector.Add(UnsupportedElementKind.LegacyDrawing, "Sheet1", null, null, "コメント");
            collector.Add(UnsupportedElementKind.LegacyDrawing, "Sheet2", null, null, "コメント");

            Assert.False(collector.IsTruncated);
        }

        [Fact]
        public void 補足の情報は正規化して入れる()
        {
            var collector = new UnsupportedElementCollector();
            collector.Add(UnsupportedElementKind.UnsupportedImageFormat, "Sheet1", null, "image/x-emf\r\n\u0000" + new string('a', 100), "画像");

            var detail = collector.Elements.Single().Detail!;
            Assert.StartsWith("image/x-emf", detail);
            Assert.DoesNotContain('\r', detail);
            Assert.DoesNotContain('\n', detail);
            Assert.DoesNotContain('\0', detail);
            Assert.Equal(DiagnosticText.MaxLength + 1, detail.Length);
            Assert.EndsWith("…", detail);
        }

        [Fact]
        public void 正規化は制御文字と書式文字と孤立したサロゲートを除く()
        {
            Assert.Equal("ab𠮷c", DiagnosticText.Normalize("a\tb​𠮷\uD800c\uDC00"));
            Assert.Null(DiagnosticText.Normalize(null));
            Assert.Equal(string.Empty, DiagnosticText.Normalize(string.Empty));
        }

        [Fact]
        public void 正規化は行区切りと段落区切りを除く()
        {
            Assert.Equal("abc", DiagnosticText.Normalize("a\u2028b\u2029c"));
        }

        [Fact]
        public void 正規化は上限ちょうどの長さなら切り詰めない()
        {
            var text = new string('x', DiagnosticText.MaxLength);
            Assert.Equal(text, DiagnosticText.Normalize(text));
        }

        [Fact]
        public void 正規化はサロゲートペアを途中で切らない()
        {
            var text = new string('x', DiagnosticText.MaxLength - 1) + "𠮷";
            var normalized = DiagnosticText.Normalize(text)!;
            Assert.Equal(new string('x', DiagnosticText.MaxLength - 1) + "…", normalized);
        }

        [Fact]
        public void 一行の説明にシートと位置を含める()
        {
            var element = new UnsupportedElement(UnsupportedElementKind.Chart, "注文書", CellAddress.Parse("D10"), null, "グラフは出力されません");
            Assert.Equal("[Chart] シート '注文書' D10: グラフは出力されません", element.ToString());

            var workbookLevel = new UnsupportedElement(UnsupportedElementKind.ExternalReference, null, null, null, "外部参照");
            Assert.Equal("[ExternalReference] ブック: 外部参照", workbookLevel.ToString());
        }

        [Fact]
        public void 未対応要素の無い例外は作れない()
        {
            Assert.Throws<System.ArgumentException>(
                () => new UnsupportedWorkbookElementException(System.Array.Empty<UnsupportedElement>(), isTruncated: false));
        }

        [Fact]
        public void 例外のメッセージは先頭の20件を載せ残りは件数だけを示す()
        {
            var elements = Enumerable.Range(1, 25)
                .Select(i => new UnsupportedElement(UnsupportedElementKind.Chart, "S", new CellAddress(i, 1), null, "グラフ" + i))
                .ToList();

            var ex = new UnsupportedWorkbookElementException(elements, isTruncated: false, reportCode: "doc");

            Assert.StartsWith("未対応の要素が 25 件あるため中止します", ex.Message);
            Assert.Contains("A20: グラフ20", ex.Message);
            Assert.DoesNotContain("グラフ21", ex.Message);
            Assert.Contains("ほか 5 件", ex.Message);
            Assert.Equal(25, ex.Elements.Count);
        }

        [Fact]
        public void 例外の診断用の文字列は先頭の100件と種類ごとの件数を載せる()
        {
            var elements = Enumerable.Range(1, 150)
                .Select(i => new UnsupportedElement(
                    i % 2 == 0 ? UnsupportedElementKind.Chart : UnsupportedElementKind.LegacyDrawing, "S", new CellAddress(i, 1), null, "要素" + i))
                .ToList();

            var diagnostic = new UnsupportedWorkbookElementException(elements, isTruncated: false).ToDiagnosticString();

            Assert.Contains("A100: 要素100", diagnostic);
            Assert.DoesNotContain("要素101", diagnostic);
            Assert.Contains("ほか 50 件", diagnostic);
            Assert.Contains("Chart=75", diagnostic);
            Assert.Contains("LegacyDrawing=75", diagnostic);
        }

        [Fact]
        public void 一行の説明はファイル由来の値を正規化し補足の情報も載せる()
        {
            var element = new UnsupportedElement(
                UnsupportedElementKind.UnsupportedImageFormat, "S\nX\r" + new string('A', 200), null, "image/x-emf", "画像");

            var text = element.ToString();

            Assert.DoesNotContain("\n", text);
            Assert.DoesNotContain("\r", text);
            Assert.True(text.Length < 200);
            Assert.EndsWith("(image/x-emf)", text);
        }

        [Fact]
        public void 例外の診断用の文字列には全件と打ち切りの印を載せる()
        {
            var elements = Enumerable.Range(1, 25)
                .Select(i => new UnsupportedElement(UnsupportedElementKind.Chart, "S", new CellAddress(i, 1), null, "グラフ" + i))
                .ToList();

            var diagnostic = new UnsupportedWorkbookElementException(elements, isTruncated: true).ToDiagnosticString();

            Assert.Contains("A25: グラフ25", diagnostic);
            Assert.Contains("件数の上限で打ち切り", diagnostic);
            Assert.Contains("25 件以上", diagnostic);
        }
    }
}
