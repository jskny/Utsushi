using System;
using System.IO;
using Xunit;
using Options = Utsushi.Cli.Program.CommandLineOptions;

namespace Utsushi.Cli.Tests
{
    /// <summary>
    /// CLI の引数解釈(<see cref="Options.Parse"/>)の検証。
    /// 帳票定義なしの変換(要件12.7)で <c>--report</c> を省略できること、定義なしでは使えない <c>--set</c>・<c>--definitions</c>
    /// の組み合わせを使い方エラーにすること、<c>--max-digit-width</c>(要件12.3)を定義なしでのみ受け付けることを中心に確かめる。
    /// </summary>
    public sealed class CommandLineOptionsTests
    {
        private const string Input = "in.xlsx";
        private const string Output = "out.pdf";

        private static Options Parse(params string[] args)
        {
            var options = Options.Parse(args, out var error);
            Assert.Null(error);
            return Assert.IsType<Options>(options);
        }

        private static string ParseError(params string[] args)
        {
            var options = Options.Parse(args, out var error);
            Assert.Null(options);
            return Assert.IsType<string>(error);
        }

        private static string[] WithDefinitionArgs(params string[] extra) =>
            Concat(new[] { "--report", "invoice", "--input", Input, "--output", Output }, extra);

        private static string[] WithoutDefinitionArgs(params string[] extra) =>
            Concat(new[] { "--input", Input, "--output", Output }, extra);

        private static string[] Concat(string[] first, string[] second)
        {
            var result = new string[first.Length + second.Length];
            first.CopyTo(result, 0);
            second.CopyTo(result, first.Length);
            return result;
        }

        // -- ヘルプ -----------------------------------------------------------------------------

        [Theory]
        [InlineData(new object[] { new string[0] })]
        [InlineData(new object[] { new[] { "--help" } })]
        [InlineData(new object[] { new[] { "-h" } })]
        [InlineData(new object[] { new[] { "--input", "in.xlsx", "--help" } })]
        [InlineData(new object[] { new[] { "--report", "invoice", "-h", "--input", "in.xlsx" } })]
        public void ヘルプ要求と引数なしはoptionsもerrorもnull(string[] args)
        {
            var options = Options.Parse(args, out var error);

            Assert.Null(options);
            Assert.Null(error);
        }

        // -- --report ---------------------------------------------------------------------------

        [Fact]
        public void reportを省略すると帳票コードはnullになり定義なしの変換になる()
        {
            var options = Parse(WithoutDefinitionArgs());

            Assert.Null(options.ReportCode);
            Assert.Equal(Input, options.InputPath);
            Assert.Equal(Output, options.OutputPath);
            Assert.Empty(options.Values);
            Assert.Empty(options.CellOverrides);
            Assert.Null(options.MaxDigitWidthPx);
        }

        [Theory]
        [InlineData("--report")]
        [InlineData("-r")]
        public void reportを指定すると帳票コードになる(string option)
        {
            var options = Parse(option, "invoice", "--input", Input, "--output", Output);

            Assert.Equal("invoice", options.ReportCode);
        }

        [Theory]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("\t")]
        public void reportの値が空ならエラー(string reportCode)
        {
            var error = ParseError("--report", reportCode, "--input", Input, "--output", Output);

            Assert.Contains("--report", error, StringComparison.Ordinal);
        }

        [Fact]
        public void reportの値が無ければエラー()
        {
            var error = ParseError("--input", Input, "--output", Output, "--report");

            Assert.Contains("--report には値が必要です", error, StringComparison.Ordinal);
        }

        // -- --set ------------------------------------------------------------------------------

        [Theory]
        [InlineData("--set")]
        [InlineData("-s")]
        public void reportなしのsetはエラー(string option)
        {
            var error = ParseError(WithoutDefinitionArgs(option, "InvoiceNo=A-001"));

            Assert.Contains("--set", error, StringComparison.Ordinal);
            Assert.Contains("--report", error, StringComparison.Ordinal);
        }

        [Fact]
        public void reportありのsetは置換キーと値になる()
        {
            var options = Parse(WithDefinitionArgs(
                "--set", "InvoiceNo=A-001",
                "-s", "Formula=a=b",          // 値の中の '=' はそのまま
                "--set", "Remarks=",          // 空の値
                "--set", "InvoiceNo=A-002")); // 同じキーは後勝ち

            Assert.Equal(3, options.Values.Count);
            Assert.Equal("A-002", options.Values["InvoiceNo"]);
            Assert.Equal("a=b", options.Values["Formula"]);
            Assert.Equal(string.Empty, options.Values["Remarks"]);
        }

        [Theory]
        [InlineData("InvoiceNo")]
        [InlineData("=A-001")]
        [InlineData("")]
        public void setがキーと値の形式でなければエラー(string pair)
        {
            var error = ParseError(WithDefinitionArgs("--set", pair));

            Assert.Contains("--set", error, StringComparison.Ordinal);
        }

        // -- --definitions ----------------------------------------------------------------------

        [Theory]
        [InlineData("--definitions")]
        [InlineData("-d")]
        public void reportなしのdefinitionsはエラー(string option)
        {
            var error = ParseError(WithoutDefinitionArgs(option, "samples/reports"));

            Assert.Contains("--definitions", error, StringComparison.Ordinal);
            Assert.Contains("--report", error, StringComparison.Ordinal);
        }

        [Fact]
        public void reportありのdefinitionsは定義ルートになる()
        {
            var options = Parse(WithDefinitionArgs("--definitions", "samples/reports"));

            Assert.Equal("samples/reports", options.DefinitionRoot);
        }

        [Fact]
        public void definitionsを省略すると定義ルートはカレントディレクトリのreportsになる()
        {
            var options = Parse(WithDefinitionArgs());

            Assert.Equal(Path.Combine(Directory.GetCurrentDirectory(), "reports"), options.DefinitionRoot);
        }

        // -- --max-digit-width ------------------------------------------------------------------

        [Theory]
        [InlineData("8", 8.0)]
        [InlineData("7", 7.0)]
        [InlineData("7.5", 7.5)]
        [InlineData("0.5", 0.5)]
        [InlineData("1e1", 10.0)]
        public void reportなしのmaxDigitWidthは正の数を受け付ける(string text, double expected)
        {
            var options = Parse(WithoutDefinitionArgs("--max-digit-width", text));

            Assert.Null(options.ReportCode);
            Assert.Equal(expected, options.MaxDigitWidthPx);
        }

        [Fact]
        public void reportありのmaxDigitWidthはエラー()
        {
            var error = ParseError(WithDefinitionArgs("--max-digit-width", "8"));

            Assert.Contains("--max-digit-width", error, StringComparison.Ordinal);
            Assert.Contains("maxDigitWidthPx", error, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData("0")]
        [InlineData("0.0")]
        [InlineData("-0")]
        [InlineData("-1")]
        [InlineData("-7.5")]
        [InlineData("abc")]
        [InlineData("")]
        [InlineData("8px")]
        [InlineData("7,5")]        // 小数点はカルチャによらず '.'
        [InlineData("Infinity")]
        [InlineData("NaN")]        // 回帰テスト: double.TryParse は NaN を受け付けるため、以前はファサードで未処理例外になった
        [InlineData("1e400")]      // double の範囲を超えて無限大になる
        public void maxDigitWidthが0や負や数値でなければエラー(string text)
        {
            var error = ParseError(WithoutDefinitionArgs("--max-digit-width", text));

            Assert.Contains("--max-digit-width には正の数を指定してください", error, StringComparison.Ordinal);
        }

        [Fact]
        public void maxDigitWidthの値が無ければエラー()
        {
            var error = ParseError(WithoutDefinitionArgs("--max-digit-width"));

            Assert.Contains("--max-digit-width には値が必要です", error, StringComparison.Ordinal);
        }

        // -- --override -------------------------------------------------------------------------

        [Fact]
        public void overrideはセル番地と値になる()
        {
            var options = Parse(WithoutDefinitionArgs(
                "--override", "A1=請求書(控)",
                "--override", "B12=x=y",       // 値の中の '=' はそのまま
                "--override", "C3=",           // 空の値
                "--override", "A1=請求書"));   // 同じ番地は後勝ち

            Assert.Equal(3, options.CellOverrides.Count);
            Assert.Equal("請求書", options.CellOverrides["A1"]);
            Assert.Equal("x=y", options.CellOverrides["B12"]);
            Assert.Equal(string.Empty, options.CellOverrides["C3"]);
            Assert.Empty(options.Values);
        }

        [Fact]
        public void overrideはreportありでも使える()
        {
            var options = Parse(WithDefinitionArgs("--override", "A1=請求書(控)", "--set", "InvoiceNo=A-001"));

            Assert.Equal("請求書(控)", options.CellOverrides["A1"]);
            Assert.Equal("A-001", options.Values["InvoiceNo"]);
        }

        [Fact]
        public void overrideのセル番地は大文字小文字を区別してそのまま渡す()
        {
            // 番地の正規化・検証は Substitution レイヤーの責務。CLI は文字列のまま渡す。
            var options = Parse(WithoutDefinitionArgs("--override", "a1=x", "--override", "A1=y"));

            Assert.Equal("x", options.CellOverrides["a1"]);
            Assert.Equal("y", options.CellOverrides["A1"]);
        }

        [Theory]
        [InlineData("A1")]
        [InlineData("=値")]
        [InlineData("")]
        public void overrideがセル番地と値の形式でなければエラー(string pair)
        {
            var error = ParseError(WithoutDefinitionArgs("--override", pair));

            Assert.Contains("--override", error, StringComparison.Ordinal);
            Assert.Contains("<セル番地>=<値>", error, StringComparison.Ordinal);
        }

        [Fact]
        public void overrideの値が無ければエラー()
        {
            var error = ParseError(WithoutDefinitionArgs("--override"));

            Assert.Contains("--override には値が必要です", error, StringComparison.Ordinal);
        }

        // -- 必須の --input / --output ---------------------------------------------------------------

        [Theory]
        [InlineData("--input", "--output")]
        [InlineData("-i", "-o")]
        [InlineData("-i", "--output")]
        public void inputとoutputは短い形式でも指定できる(string inputOption, string outputOption)
        {
            var options = Parse(inputOption, Input, outputOption, Output);

            Assert.Equal(Input, options.InputPath);
            Assert.Equal(Output, options.OutputPath);
        }

        [Fact]
        public void inputが無ければエラー()
        {
            Assert.Equal("--input は必須です。", ParseError("--output", Output));
            Assert.Equal("--input は必須です。", ParseError("--report", "invoice", "--output", Output));
        }

        [Fact]
        public void outputが無ければエラー()
        {
            Assert.Equal("--output は必須です。", ParseError("--input", Input));
            Assert.Equal("--output は必須です。", ParseError("--report", "invoice", "--input", Input));
        }

        [Theory]
        [InlineData("")]
        [InlineData("  ")]
        public void inputとoutputの値が空ならエラー(string value)
        {
            Assert.Equal("--input は必須です。", ParseError("--input", value, "--output", Output));
            Assert.Equal("--output は必須です。", ParseError("--input", Input, "--output", value));
        }

        [Fact]
        public void inputの値が無ければエラー()
        {
            Assert.Contains("--input には値が必要です", ParseError("--output", Output, "--input"), StringComparison.Ordinal);
        }

        // -- その他のオプション ------------------------------------------------------------------------

        [Fact]
        public void 不明なオプションはエラー()
        {
            var error = ParseError(WithoutDefinitionArgs("--unknown"));

            Assert.Contains("--unknown", error, StringComparison.Ordinal);
        }

        [Fact]
        public void 出力オプションは既定でオフ()
        {
            var options = Parse(WithoutDefinitionArgs());

            Assert.False(options.AllowFontFallback);
            Assert.Null(options.FallbackFont);
            Assert.False(options.AllowMissingGlyphs);
            Assert.False(options.OutlineText);
        }

        [Fact]
        public void 出力オプションは定義なしでも使える()
        {
            var options = Parse(WithoutDefinitionArgs("--allow-font-fallback", "--allow-missing-glyphs", "--outline-text"));

            Assert.True(options.AllowFontFallback);
            Assert.Null(options.FallbackFont);
            Assert.True(options.AllowMissingGlyphs);
            Assert.True(options.OutlineText);
        }

        [Fact]
        public void allowFontFallbackの直後がオプションでなければ代替フォント名になる()
        {
            var options = Parse(WithoutDefinitionArgs("--allow-font-fallback", "IPAexGothic", "--outline-text"));

            Assert.True(options.AllowFontFallback);
            Assert.Equal("IPAexGothic", options.FallbackFont);
            Assert.True(options.OutlineText);
        }
    }
}
