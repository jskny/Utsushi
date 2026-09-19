using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Utsushi;
using Utsushi.Core.Exceptions;
using Utsushi.Layout;
using Utsushi.Layout.Model;
using Utsushi.Layout.Text;
using Utsushi.Parsing.OpenXml;
using Utsushi.Rendering;
using Utsushi.ReportDefinitions;
using Utsushi.Substitution;
using Xunit;

namespace Utsushi.Golden.Tests;

/// <summary>
/// 登録済み帳票サンプルのエンドツーエンド(Parsing → Rendering)ゴールデンテスト(タスク8.2、要件8.3)。
/// </summary>
/// <remarks>
/// <para>
/// レイアウトの比較には <see cref="ApproximateFontMetricsProvider"/> を使う。
/// 実フォントのメトリクスは実行環境のフォント構成に依存し、CIとローカルでゴールデンが
/// 一致しなくなるため、回帰検出の基準としては決定的な近似メトリクスを用いる。
/// </para>
/// <para>
/// 実フォントを使った見た目の確認は「PDFが生成できること」までを自動テストの範囲とし、
/// 最終的な見た目のレビューは人が行う前提とする。
/// </para>
/// </remarks>
public sealed class ReportConversionGoldenTests
{
    /// <summary>
    /// 決定的なフォントメトリクスを使い、PDF描画の手前までを実行するコンバータ。
    /// </summary>
    private static ReportPdfConverter CreateDeterministicConverter(out FontResolver fontResolver)
    {
        fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
        var skiaMetrics = new SkiaFontMetricsProvider(fontResolver);

        return new ReportPdfConverter(
            new OpenXmlWorkbookReader(),
            new FileSystemReportDefinitionRepository(TestPaths.SampleReportsRoot),
            new ReportModelBuilder(),
            new CellSubstitutor(),
            new ReportLayoutEngine(new ApproximateFontMetricsProvider()),
            new SkiaPdfRenderer(skiaMetrics),
            fontResolver);
    }

    public static IEnumerable<object[]> RegisteredReports()
    {
        yield return new object[]
        {
            "invoice",
            new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 御中",
                ["InvoiceNo"] = "INV-2026-0417",
                ["Subject"] = "件名: ゴールデンテスト用案件",
                ["TotalAmount"] = "¥2,153,800",
                ["Remarks"] = "備考:\nゴールデンテストで固定した値です。",
            },
        };

        yield return new object[]
        {
            "delivery-note",
            new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 御中",
                ["DeliveryNo"] = "DN-2026-0088",
                ["DeliveryDate"] = "2026年4月15日",
            },
        };
    }

    [Theory]
    [MemberData(nameof(RegisteredReports))]
    public void 登録済み帳票のレイアウトが期待結果と一致する(string reportCode, Dictionary<string, string> values)
    {
        using var converter = CreateDeterministicConverter(out _);
        using var input = File.OpenRead(TestPaths.SampleTemplate(reportCode));

        var layout = converter.ComputeLayout(reportCode, input, values);

        GoldenFile.Verify(reportCode, "layout.snapshot.txt", GoldenSnapshot.Create(layout));
    }

    [Theory]
    [MemberData(nameof(RegisteredReports))]
    public void 登録済み帳票からPDFを生成できる(string reportCode, Dictionary<string, string> values)
    {
        using var converter = CreateDeterministicConverter(out _);
        using var input = File.OpenRead(TestPaths.SampleTemplate(reportCode));
        using var output = new MemoryStream();

        converter.Convert(reportCode, input, values, output);

        var bytes = output.ToArray();
        Assert.True(bytes.Length > 0, "PDFが出力されるはず");
        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
    }

    [Fact]
    public void 請求書は1ページに収まる()
    {
        var layout = ComputeInvoiceLayout();

        Assert.Equal(1, layout.PageCount);
        Assert.Equal("A4", layout.Pages[0].Paper.Name);
    }

    [Fact]
    public void 納品書は複数ページになり印刷タイトルが各ページに繰り返される()
    {
        using var converter = CreateDeterministicConverter(out _);
        using var input = File.OpenRead(TestPaths.SampleTemplate("delivery-note"));

        var layout = converter.ComputeLayout(
            "delivery-note",
            input,
            new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 御中",
                ["DeliveryNo"] = "DN-2026-0088",
            });

        Assert.True(layout.PageCount > 1, $"複数ページになるはず (実際: {layout.PageCount})");

        // 印刷タイトル行(1〜7行目)の内容が全ページに現れる(要件3.4)
        foreach (var page in layout.Pages)
        {
            var texts = page.Commands.OfType<TextCommand>().Select(t => t.Text).ToList();
            Assert.Contains("納品書", texts);
            Assert.Contains("品名", texts);
            Assert.Contains("株式会社テスト製作所 御中", texts);
        }

        // 本文はページをまたいで連続する
        var firstPageRows = layout.Pages[0].RowRange;
        var secondPageRows = layout.Pages[1].RowRange;
        Assert.Equal(firstPageRows.Last + 1, secondPageRows.First);
    }

    [Fact]
    public void 置換値が帳票に反映される()
    {
        var layout = ComputeInvoiceLayout();
        var texts = layout.Pages
            .SelectMany(p => p.Commands.OfType<TextCommand>())
            .Select(t => t.Text)
            .ToList();

        Assert.Contains("株式会社テスト製作所 御中", texts);
        Assert.Contains("INV-2026-0417", texts);
        Assert.Contains("¥2,153,800", texts);

        // テンプレートの既定値は残らない
        Assert.DoesNotContain("INV-0000", texts);
    }

    [Fact]
    public void 必須の置換値が無ければ変換前に失敗する()
    {
        using var converter = CreateDeterministicConverter(out _);
        using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

        var ex = Assert.Throws<RequiredSubstitutionValueMissingException>(
            () => converter.ComputeLayout("invoice", input, new Dictionary<string, string>()));

        Assert.Equal("invoice", ex.ReportCode);
    }

    [Fact]
    public void 未登録の帳票コードは明確なエラーになる()
    {
        using var converter = CreateDeterministicConverter(out _);
        using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

        var ex = Assert.Throws<ReportDefinitionNotFoundException>(
            () => converter.ComputeLayout("未登録の帳票", input, new Dictionary<string, string>()));

        Assert.Equal("未登録の帳票", ex.ReportCode);
    }

    [Fact]
    public void 帳票定義と一致しないExcelを渡すと失敗する()
    {
        // 納品書のテンプレートを請求書の定義で読もうとするとシート名が一致しない(要件1.4)
        using var converter = CreateDeterministicConverter(out _);
        using var input = File.OpenRead(TestPaths.SampleTemplate("delivery-note"));

        var ex = Assert.Throws<InvalidExcelFileException>(
            () => converter.ComputeLayout(
                "invoice",
                input,
                new Dictionary<string, string>
                {
                    ["CustomerName"] = "x",
                    ["InvoiceNo"] = "y",
                    ["TotalAmount"] = "z",
                }));

        Assert.Equal(InvalidExcelFileReason.NoWorksheet, ex.Reason);
        Assert.Equal("invoice", ex.ReportCode);
    }

    private static PagedLayout ComputeInvoiceLayout()
    {
        using var converter = CreateDeterministicConverter(out _);
        using var input = File.OpenRead(TestPaths.SampleTemplate("invoice"));

        return converter.ComputeLayout(
            "invoice",
            input,
            new Dictionary<string, string>
            {
                ["CustomerName"] = "株式会社テスト製作所 御中",
                ["InvoiceNo"] = "INV-2026-0417",
                ["TotalAmount"] = "¥2,153,800",
            });
    }
}
