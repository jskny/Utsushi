using Utsushi.Core;
using Utsushi.Rendering;
using Xunit;

namespace Utsushi.Cli.Tests
{
    /// <summary>
    /// CLI の引数解釈(<see cref="CommandLineOptions.Parse"/>)の検証(excel2pdf-consolidation 要件5.4〜5.6・5.8)。
    /// </summary>
    public sealed class CommandLineOptionsTests
    {
        private const string Input = "in.xlsx";
        private const string Output = "out.pdf";

        private static CommandLineOptions Parse(params string[] args)
        {
            var options = CommandLineOptions.Parse(args, out var error);
            Assert.Null(error);
            return Assert.IsType<CommandLineOptions>(options);
        }

        private static string ParseError(params string[] args)
        {
            var options = CommandLineOptions.Parse(args, out var error);
            Assert.Null(options);
            return Assert.IsType<string>(error);
        }

        [Fact]
        public void checkは入力のパスだけを受け取る()
        {
            var options = Parse("check", Input);

            Assert.Equal(CliCommand.Check, options.Command);
            Assert.Equal(Input, options.InputPath);
            Assert.Null(options.OutputPath);
        }

        [Fact]
        public void convertは入力と出力のパスを受け取る()
        {
            var options = Parse("convert", Input, Output);

            Assert.Equal(CliCommand.Convert, options.Command);
            Assert.Equal(Input, options.InputPath);
            Assert.Equal(Output, options.OutputPath);
        }

        [Theory]
        [InlineData("--help")]
        [InlineData("-h")]
        public void ヘルプはどの位置に書いても使い方の表示にする(string help)
        {
            Assert.Null(CommandLineOptions.Parse(new[] { help }, out var error));
            Assert.Null(error);
            Assert.Null(CommandLineOptions.Parse(new[] { "convert", Input, Output, "--text", "A1=x", help }, out error));
            Assert.Null(error);
        }

        [Fact]
        public void コマンドが無ければ使い方エラーにする()
        {
            Assert.Contains("コマンド", ParseError());
        }

        [Fact]
        public void 不明なコマンドは使い方エラーにする()
        {
            Assert.Contains("'render'", ParseError("render", Input));
        }

        [Theory]
        [InlineData("check")]
        [InlineData("convert", Input)]
        [InlineData("convert", Input, "--text", "A1=x")]
        public void 必須のパスが無ければ使い方エラーにする(params string[] args)
        {
            Assert.Contains("パス", ParseError(args));
        }

        [Fact]
        public void 余分なパスは使い方エラーにする()
        {
            Assert.Contains("'extra.pdf'", ParseError("check", Input, "extra.pdf"));
        }

        [Fact]
        public void オプションの後ろのパスは使い方エラーにする()
        {
            // --allow-font-fallback の書体名と取り違えないよう、パスはコマンドの直後に限る。
            Assert.Contains("コマンドの直後", ParseError("check", Input, "--ignore-unsupported", "other.xlsx"));
        }

        [Fact]
        public void NUL文字を含むパスは使い方エラーにする()
        {
            Assert.Contains("NUL", ParseError("convert", Input, "out\0.pdf"));
        }

        [Theory]
        [InlineData("--report")]
        [InlineData("--definitions")]
        [InlineData("--set")]
        [InlineData("--override")]
        [InlineData("--input")]
        public void 帳票定義と以前の形式のオプションは持たない(string option)
        {
            Assert.Contains("不明なオプション", ParseError("convert", Input, Output, option, "x"));
        }

        [Fact]
        public void 文字列と数値のセルの指定を受け取る()
        {
            var options = Parse(
                "convert", Input, Output,
                "--text", "A3=株式会社サンプル 御中",
                "--text", "B4=式は=を含められる",
                "--number", "F12=-1234.5",
                "--number", "c8=1e3");

            Assert.Equal("株式会社サンプル 御中", options.Texts[CellAddress.Parse("A3")]);
            Assert.Equal("式は=を含められる", options.Texts[CellAddress.Parse("B4")]);
            Assert.Equal(-1234.5, options.Numbers[CellAddress.Parse("F12")]);
            Assert.Equal(1000.0, options.Numbers[CellAddress.Parse("C8")]);
        }

        [Fact]
        public void 空の文字列を書き込める()
        {
            Assert.Equal(string.Empty, Parse("check", Input, "--text", "A1=").Texts[CellAddress.Parse("A1")]);
        }

        [Theory]
        [InlineData("--text", "A1=x", "--text", "A1=y")]
        [InlineData("--text", "A1=x", "--number", "a1=1")]
        [InlineData("--number", "$B$2=1", "--text", "B2=x")]
        public void 同じセルを重複して指定すると使い方エラーにする(string option1, string value1, string option2, string value2)
        {
            Assert.Contains("重複", ParseError("check", Input, option1, value1, option2, value2));
        }

        [Theory]
        [InlineData("--text", "A1")]
        [InlineData("--text", "=x")]
        [InlineData("--text", "1A=x")]
        [InlineData("--number", "A1=1,5")]
        [InlineData("--number", "A1=NaN")]
        [InlineData("--number", "A1=Infinity")]
        [InlineData("--number", "A1=")]
        public void セルの指定の形式が不正なら使い方エラーにする(string option, string value)
        {
            Assert.Contains(option, ParseError("check", Input, option, value));
        }

        [Theory]
        [InlineData("--text")]
        [InlineData("--number")]
        [InlineData("--sheet")]
        [InlineData("--document-name")]
        [InlineData("--max-digit-width")]
        public void 値の必要なオプションに値が無ければ使い方エラーにする(string option)
        {
            Assert.Contains("値が必要", ParseError("check", Input, option));
        }

        [Theory]
        [InlineData("--sheet")]
        [InlineData("--document-name")]
        public void 空の名前は使い方エラーにする(string option)
        {
            Assert.Contains("空", ParseError("check", Input, option, " "));
        }

        [Theory]
        [InlineData("0")]
        [InlineData("-7")]
        [InlineData("abc")]
        [InlineData("NaN")]
        public void 最大数字幅は正の数だけを受け付ける(string value)
        {
            Assert.Contains("--max-digit-width", ParseError("check", Input, "--max-digit-width", value));
        }

        [Fact]
        public void 設定のオプションをExcel2PdfOptionsに渡す()
        {
            var options = Parse(
                "convert", Input, Output,
                "--sheet", "請求書",
                "--document-name", "INV-0001",
                "--max-digit-width", "7.5",
                "--ignore-unsupported",
                "--allow-font-fallback", "IPAexGothic",
                "--allow-missing-glyphs",
                "--outline-text").ToExcel2PdfOptions();

            Assert.Equal("請求書", options.SheetName);
            Assert.Equal("INV-0001", options.DocumentName);
            Assert.Equal(7.5, options.MaxDigitWidthPx);
            Assert.Equal(UnsupportedElementPolicy.Ignore, options.UnsupportedElements);
            Assert.Equal(FontResolutionMode.AllowFallback, options.Fonts!.Mode);
            Assert.Equal("IPAexGothic", options.Fonts.FallbackFamilyName);
            Assert.Equal(MissingGlyphPolicy.Render, options.Render!.MissingGlyphs);
            Assert.Equal(PdfTextRendering.Outline, options.Render.TextRendering);
        }

        [Fact]
        public void 設定のオプションを省略すると既定の設定にする()
        {
            var options = Parse("check", Input).ToExcel2PdfOptions();

            Assert.Null(options.SheetName);
            Assert.Null(options.DocumentName);
            Assert.Null(options.MaxDigitWidthPx);
            Assert.Equal(UnsupportedElementPolicy.Error, options.UnsupportedElements);
            Assert.Equal(FontResolutionMode.Strict, options.Fonts!.Mode);
            Assert.Equal(PdfRenderOptions.Default, options.Render);
        }

        [Theory]
        [InlineData("--sheet", "a", "--sheet", "b")]
        [InlineData("--ignore-unsupported", null, "--ignore-unsupported", null)]
        public void テキストと数値以外のオプションを重ねて指定すると使い方エラーにする(string first, string? firstValue, string second, string? secondValue)
        {
            var args = new System.Collections.Generic.List<string> { "check", Input, first };
            if (firstValue is not null)
            {
                args.Add(firstValue);
            }

            args.Add(second);
            if (secondValue is not null)
            {
                args.Add(secondValue);
            }

            Assert.Contains("重複", ParseError(args.ToArray()));
        }

        [Fact]
        public void 値の代わりに次のオプションを書くと使い方エラーにする()
        {
            Assert.Contains("値が必要", ParseError("check", Input, "--document-name", "--allow-font-fallback"));
        }

        [Fact]
        public void 形式の誤りのエラーに書き込む値を出さない()
        {
            Assert.DoesNotContain("株式会社", ParseError("check", Input, "--text", "株式会社サンプル"));
        }

        [Fact]
        public void 書体名を省略したフォントの代替の許可は直後のオプションを書体名にしない()
        {
            var options = Parse("check", Input, "--allow-font-fallback", "--outline-text");

            Assert.True(options.AllowFontFallback);
            Assert.Null(options.FallbackFont);
            Assert.True(options.OutlineText);
        }
    }
}
