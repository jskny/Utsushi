using System;
using System.IO;
using System.Text;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.TestSupport;
using Xunit;

namespace Utsushi.Cli.Tests
{
    /// <summary>
    /// <see cref="Program.Run"/> の終了コードと出力の検証(excel2pdf-consolidation 要件5.1〜5.3・5.6・5.7)。
    /// </summary>
    /// <remarks>
    /// 開発環境・CIにサンプルの書体(MS PGothic・游ゴシック)が無いことがあるため、変換は <c>--allow-font-fallback</c> を付けて行う。
    /// その場合 <c>check</c> はフォントの代替を問題として数えるため、終了コード0の判定は <see cref="Program.CheckExitCode"/> で確かめる。
    /// </remarks>
    public sealed class ProgramRunTests : IDisposable
    {
        private static readonly string Invoice = TestPaths.SampleTemplate("invoice");
        private static readonly string OrderForm = Path.Combine(TestPaths.EdgeCaseSamplesRoot, "edge-order-form", "template.xlsx");

        private readonly string _directory = Path.Combine(Path.GetTempPath(), "utsushi-cli-test-" + Guid.NewGuid().ToString("N"));

        public ProgramRunTests()
        {
            Directory.CreateDirectory(_directory);
        }

        public void Dispose()
        {
            Directory.Delete(_directory, recursive: true);
        }

        private static (int ExitCode, string Stdout, string Stderr) Run(params string[] args)
        {
            using var stdout = new StringWriter();
            using var stderr = new StringWriter();
            var exitCode = Program.Run(args, stdout, stderr);
            return (exitCode, stdout.ToString(), stderr.ToString());
        }

        [Fact]
        public void ヘルプは使い方を標準出力に出して終了コード0にする()
        {
            var (exitCode, stdout, stderr) = Run("--help");

            Assert.Equal(Program.ExitSuccess, exitCode);
            Assert.Contains("utsushi check", stdout, StringComparison.Ordinal);
            Assert.Contains("utsushi convert", stdout, StringComparison.Ordinal);
            Assert.Empty(stderr);
        }

        [Fact]
        public void 使い方エラーは理由と使い方を標準エラーに出して終了コード2にする()
        {
            var (exitCode, stdout, stderr) = Run("convert", Invoice);

            Assert.Equal(Program.ExitUsageError, exitCode);
            Assert.Empty(stdout);
            Assert.Contains("エラー:", stderr, StringComparison.Ordinal);
            Assert.Contains("utsushi convert", stderr, StringComparison.Ordinal);
        }

        [Fact]
        public void convertは書き込んだ値でPDFを保存して終了コード0にする()
        {
            var output = Path.Combine(_directory, "invoice.pdf");

            var (exitCode, stdout, stderr) = Run(
                "convert", Invoice, output, "--text", "A3=株式会社テスト 御中", "--number", "F12=1234", "--allow-font-fallback");

            Assert.Equal(Program.ExitSuccess, exitCode);
            Assert.Contains(output, stdout, StringComparison.Ordinal);
            Assert.Empty(stderr);
            Assert.Equal("%PDF", Encoding.ASCII.GetString(File.ReadAllBytes(output), 0, 4));
        }

        [Fact]
        public void 未対応要素で変換を中止すると要素の一覧を標準エラーに出して終了コード3にする()
        {
            var output = Path.Combine(_directory, "order.pdf");

            var (exitCode, _, stderr) = Run("convert", OrderForm, output, "--allow-font-fallback");

            Assert.Equal(Program.ExitConversionError, exitCode);
            Assert.Contains("変換に失敗しました。", stderr, StringComparison.Ordinal);
            Assert.Contains("[Chart]", stderr, StringComparison.Ordinal);
            Assert.Contains("[LegacyDrawing]", stderr, StringComparison.Ordinal);
            Assert.False(File.Exists(output));
        }

        [Fact]
        public void 未対応要素を無視する指定なら変換できる()
        {
            var output = Path.Combine(_directory, "order.pdf");

            var (exitCode, _, _) = Run("convert", OrderForm, output, "--allow-font-fallback", "--ignore-unsupported");

            Assert.Equal(Program.ExitSuccess, exitCode);
            Assert.True(File.Exists(output));
        }

        [Fact]
        public void 入力ファイルが無ければ終了コード3にする()
        {
            var (exitCode, _, stderr) = Run("check", Path.Combine(_directory, "存在しない.xlsx"), "--allow-font-fallback");

            Assert.Equal(Program.ExitConversionError, exitCode);
            Assert.Contains("診断に失敗しました。", stderr, StringComparison.Ordinal);
        }

        [Fact]
        public void 書き込む値が不正なら終了コード3にする()
        {
            var (exitCode, _, stderr) = Run("check", Invoice, "--text", "F3=INV\t0001", "--allow-font-fallback");

            Assert.Equal(Program.ExitConversionError, exitCode);
            Assert.Contains("F3", stderr, StringComparison.Ordinal);
        }

        [Fact]
        public void checkは種類ごとに件数と一覧を出し問題があれば終了コード1にする()
        {
            const string payment = "月末締め翌月末払い(プルダウンの選択肢に無い値)";

            var (exitCode, stdout, stderr) = Run(
                "check", OrderForm, "--text", "B5=" + payment, "--max-digit-width", "8", "--allow-font-fallback");

            Assert.Equal(Program.ExitIssuesFound, exitCode);
            Assert.Empty(stderr);
            Assert.Contains("未対応の要素: ", stdout, StringComparison.Ordinal);
            Assert.Contains("  [Chart] シート '注文書' ", stdout, StringComparison.Ordinal);
            Assert.Contains("文字の収まらない箇所: ", stdout, StringComparison.Ordinal);
            Assert.Contains("  [Clipped] B5(1ページ目): ", stdout, StringComparison.Ordinal);
            Assert.Contains("書式の近似(参考): ", stdout, StringComparison.Ordinal);

            // 書き込んだ文字列(個人情報を含みうる)は出さない(design.md 5.2)。
            Assert.DoesNotContain(payment, stdout, StringComparison.Ordinal);

            // 書式の近似は最後に出す。
            Assert.True(
                stdout.IndexOf("書式の近似(参考)", StringComparison.Ordinal) > stdout.IndexOf("文字の収まらない箇所", StringComparison.Ordinal));
        }

        [Fact]
        public void checkは未対応要素を無視する指定でも未対応要素を問題として数える()
        {
            var (exitCode, stdout, _) = Run("check", OrderForm, "--ignore-unsupported", "--allow-font-fallback");

            Assert.Equal(Program.ExitIssuesFound, exitCode);
            Assert.Contains("[Chart]", stdout, StringComparison.Ordinal);
        }

        [Fact]
        public void 書式の近似だけなら終了コード0にし参考として出す()
        {
            var approximation = new FormatApproximation(
                FormatApproximationKind.TextRotation, "工程表", CellRange.Parse("D3:D3"), "textRotation=90", "文字の回転は、回転せずに横書きで描きます。");
            var diagnostics = Diagnostics(approximations: new[] { approximation });
            using var output = new StringWriter();

            Program.WriteDiagnostics(diagnostics, output);

            Assert.Equal(Program.ExitSuccess, Program.CheckExitCode(diagnostics));
            Assert.Equal(
                "書式の近似(参考): 1件" + Environment.NewLine
                + "  [TextRotation] シート '工程表' D3: 文字の回転は、回転せずに横書きで描きます。 (textRotation=90)" + Environment.NewLine
                + "問題は見つかりませんでした(書式の近似は参考情報です)。" + Environment.NewLine,
                output.ToString());
        }

        [Fact]
        public void 問題が無ければ終了コード0にしその旨を出す()
        {
            var diagnostics = Diagnostics();
            using var output = new StringWriter();

            Program.WriteDiagnostics(diagnostics, output);

            Assert.Equal(Program.ExitSuccess, Program.CheckExitCode(diagnostics));
            Assert.Equal("問題は見つかりませんでした。" + Environment.NewLine, output.ToString());
        }

        [Fact]
        public void フォントの代替と打ち切りの印を出し問題として数える()
        {
            var fit = new FitIssue(FitIssueKind.NumberTooWide, CellAddress.Parse("C8"), 1, "¥123", true, "C8 の数値が収まりません。");
            var diagnostics = Diagnostics(
                fitIssues: new[] { fit },
                fitIssuesTruncated: true,
                fonts: new[] { new FontSubstitution("ＭＳ Ｐゴシック", true, false, "BIZ UDPGothic") });
            using var output = new StringWriter();

            Program.WriteDiagnostics(diagnostics, output);

            Assert.Equal(Program.ExitIssuesFound, Program.CheckExitCode(diagnostics));
            Assert.Equal(
                "文字の収まらない箇所: 1件(上限で打ち切り)" + Environment.NewLine
                + "  [NumberTooWide] C8(1ページ目): C8 の数値が収まりません。" + Environment.NewLine
                + "フォントの代替: 1件" + Environment.NewLine
                + "  ＭＳ Ｐゴシック(太字) → BIZ UDPGothic" + Environment.NewLine,
                output.ToString());
        }

        private static Excel2PdfDiagnostics Diagnostics(
            FitIssue[]? fitIssues = null,
            bool fitIssuesTruncated = false,
            FontSubstitution[]? fonts = null,
            FormatApproximation[]? approximations = null) =>
            new(
                fitIssues ?? Array.Empty<FitIssue>(),
                fitIssuesTruncated,
                Array.Empty<UnsupportedElement>(),
                false,
                fonts ?? Array.Empty<FontSubstitution>(),
                approximations ?? Array.Empty<FormatApproximation>(),
                false);
    }
}
