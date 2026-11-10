using System;
using System.Collections.Generic;
using System.Globalization;
using Utsushi.Core;
using Utsushi.Rendering;

namespace Utsushi.Cli
{
    /// <summary>CLI のコマンド(excel2pdf-consolidation 要件5)。</summary>
    internal enum CliCommand
    {
        /// <summary>診断する(<see cref="Excel2Pdf.Check"/>)。</summary>
        Check,

        /// <summary>PDF に変換して保存する(<see cref="Excel2Pdf.Save(string)"/>)。</summary>
        Convert,
    }

    /// <summary>
    /// コマンドライン引数(excel2pdf-consolidation 要件5、design.md 5.1)。
    /// </summary>
    /// <remarks>
    /// <para>形式: <c>check &lt;入力.xlsx&gt; [オプション]</c>、<c>convert &lt;入力.xlsx&gt; &lt;出力.pdf&gt; [オプション]</c>。</para>
    /// <para>
    /// 入力・出力のパスは、コマンドの直後(オプションより前)に書く。<c>--allow-font-fallback</c> の書体名は省略できるため、
    /// パスをオプションの後ろにも書けるようにすると、パスを書体名と取り違えるためである。
    /// </para>
    /// </remarks>
    internal sealed class CommandLineOptions
    {
        private CommandLineOptions(
            CliCommand command,
            string inputPath,
            string? outputPath,
            IReadOnlyDictionary<CellAddress, string> texts,
            IReadOnlyDictionary<CellAddress, double> numbers,
            string? sheetName,
            string? documentName,
            double? maxDigitWidthPx,
            bool ignoreUnsupported,
            bool allowFontFallback,
            string? fallbackFont,
            bool allowMissingGlyphs,
            bool outlineText)
        {
            Command = command;
            InputPath = inputPath;
            OutputPath = outputPath;
            Texts = texts;
            Numbers = numbers;
            SheetName = sheetName;
            DocumentName = documentName;
            MaxDigitWidthPx = maxDigitWidthPx;
            IgnoreUnsupported = ignoreUnsupported;
            AllowFontFallback = allowFontFallback;
            FallbackFont = fallbackFont;
            AllowMissingGlyphs = allowMissingGlyphs;
            OutlineText = outlineText;
        }

        public CliCommand Command { get; }

        public string InputPath { get; }

        /// <summary>出力の PDF のパス。<see cref="CliCommand.Convert"/> だけ。</summary>
        public string? OutputPath { get; }

        /// <summary><c>--text</c> で指定したセルと文字列(<see cref="Excel2Pdf.SetText"/>)。</summary>
        public IReadOnlyDictionary<CellAddress, string> Texts { get; }

        /// <summary><c>--number</c> で指定したセルと数値(<see cref="Excel2Pdf.SetValue(string, double)"/>)。</summary>
        public IReadOnlyDictionary<CellAddress, double> Numbers { get; }

        public string? SheetName { get; }

        public string? DocumentName { get; }

        public double? MaxDigitWidthPx { get; }

        public bool IgnoreUnsupported { get; }

        public bool AllowFontFallback { get; }

        /// <summary>代替の書体名。省略すると同梱のフォント。</summary>
        public string? FallbackFont { get; }

        public bool AllowMissingGlyphs { get; }

        public bool OutlineText { get; }

        /// <summary>変換の設定(要件5.5)。</summary>
        public Excel2PdfOptions ToExcel2PdfOptions()
        {
            var render = OutlineText ? PdfRenderOptions.OutlineText : PdfRenderOptions.Default;
            if (AllowMissingGlyphs)
            {
                render = render with { MissingGlyphs = MissingGlyphPolicy.Render };
            }

            return new Excel2PdfOptions
            {
                Fonts = AllowFontFallback ? FontResolverOptions.AllowFallback(FallbackFont) : FontResolverOptions.Strict,
                Render = render,
                SheetName = SheetName,
                DocumentName = DocumentName,
                MaxDigitWidthPx = MaxDigitWidthPx,
                UnsupportedElements = IgnoreUnsupported ? UnsupportedElementPolicy.Ignore : UnsupportedElementPolicy.Error,
            };
        }

        /// <summary>
        /// 引数を解釈する。使い方の表示を求められた場合(<c>--help</c>)は、options・error ともに null を返す。
        /// 引数が不正な場合は options を null、error に理由を返す(要件5.6)。
        /// </summary>
        public static CommandLineOptions? Parse(string[] args, out string? error)
        {
            error = null;
            if (Array.Exists(args, a => a is "--help" or "-h"))
            {
                return null;
            }

            if (args.Length == 0)
            {
                error = "コマンド(check または convert)を指定してください。";
                return null;
            }

            CliCommand command;
            switch (args[0])
            {
                case "check":
                    command = CliCommand.Check;
                    break;
                case "convert":
                    command = CliCommand.Convert;
                    break;
                default:
                    error = $"不明なコマンドです: '{args[0]}'(check または convert を指定してください)";
                    return null;
            }

            // コマンドの直後、最初のオプションまでがパス。
            var positionals = new List<string>();
            var i = 1;
            for (; i < args.Length && !IsOption(args[i]); i++)
            {
                positionals.Add(args[i]);
            }

            var expected = command == CliCommand.Check ? 1 : 2;
            if (positionals.Count < expected)
            {
                error = command == CliCommand.Check
                    ? "check には入力の Excel ファイルのパスを指定してください。"
                    : "convert には入力の Excel ファイルと出力の PDF のパスを指定してください。";
                return null;
            }

            if (positionals.Count > expected)
            {
                error = $"余分な引数です: '{positionals[expected]}'";
                return null;
            }

            foreach (var path in positionals)
            {
                if (string.IsNullOrWhiteSpace(path) || path.IndexOf('\0') >= 0)
                {
                    error = "パスが空か、使えない文字(NUL)を含みます。";
                    return null;
                }
            }

            var texts = new Dictionary<CellAddress, string>();
            var numbers = new Dictionary<CellAddress, double>();
            string? sheetName = null, documentName = null, fallbackFont = null;
            double? maxDigitWidthPx = null;
            bool ignoreUnsupported = false, allowFontFallback = false, allowMissingGlyphs = false, outlineText = false;

            var seenOptions = new HashSet<string>(StringComparer.Ordinal);
            for (; i < args.Length; i++)
            {
                // --text・--number 以外のオプションを重ねて指定した場合は、どちらを使うか決められないため使い方エラーにする。
                if (args[i] is not ("--text" or "--number") && IsOption(args[i]) && !seenOptions.Add(args[i]))
                {
                    error = $"{args[i]} が重複して指定されています。1つにしてください。";
                    return null;
                }

                switch (args[i])
                {
                    case "--text":
                        {
                            if (!TryTakeValue(args, ref i, out var pair, out error)
                                || !TryParseCellValue(pair!, "--text", "<セル>=<文字列>", texts, numbers, out var cell, out var text, out error))
                            {
                                return null;
                            }

                            texts.Add(cell, text);
                            break;
                        }

                    case "--number":
                        {
                            if (!TryTakeValue(args, ref i, out var pair, out error)
                                || !TryParseCellValue(pair!, "--number", "<セル>=<数値>", texts, numbers, out var cell, out var text, out error))
                            {
                                return null;
                            }

                            // 小数点は「.」。カルチャに依存しない(design.md 5.1)。桁区切りは受け付けない。
                            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                                || double.IsNaN(number) || double.IsInfinity(number))
                            {
                                error = $"--number の値は数値で指定してください(小数点は「.」): セル {cell}";
                                return null;
                            }

                            numbers.Add(cell, number);
                            break;
                        }

                    case "--sheet":
                        if (!TryTakeNonEmptyValue(args, ref i, out sheetName, out error))
                        {
                            return null;
                        }

                        break;

                    case "--document-name":
                        if (!TryTakeNonEmptyValue(args, ref i, out documentName, out error))
                        {
                            return null;
                        }

                        break;

                    case "--max-digit-width":
                        {
                            if (!TryTakeValue(args, ref i, out var text, out error))
                            {
                                return null;
                            }

                            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var width)
                                || double.IsNaN(width) || double.IsInfinity(width) || width <= 0)
                            {
                                error = $"--max-digit-width には正の数を指定してください: '{text}'";
                                return null;
                            }

                            maxDigitWidthPx = width;
                            break;
                        }

                    case "--ignore-unsupported":
                        ignoreUnsupported = true;
                        break;

                    case "--allow-font-fallback":
                        allowFontFallback = true;

                        // 直後の引数がオプションでなければ代替の書体名として扱う。
                        if (i + 1 < args.Length && !IsOption(args[i + 1]))
                        {
                            fallbackFont = args[++i];
                            if (string.IsNullOrWhiteSpace(fallbackFont))
                            {
                                error = "--allow-font-fallback の書体名が空です。";
                                return null;
                            }
                        }

                        break;

                    case "--allow-missing-glyphs":
                        allowMissingGlyphs = true;
                        break;

                    case "--outline-text":
                        outlineText = true;
                        break;

                    default:
                        error = IsOption(args[i])
                            ? $"不明なオプションです: '{args[i]}'"
                            : $"余分な引数です: '{args[i]}'(パスはコマンドの直後に書いてください)";
                        return null;
                }
            }

            return new CommandLineOptions(
                command,
                positionals[0],
                command == CliCommand.Convert ? positionals[1] : null,
                texts,
                numbers,
                sheetName,
                documentName,
                maxDigitWidthPx,
                ignoreUnsupported,
                allowFontFallback,
                fallbackFont,
                allowMissingGlyphs,
                outlineText);
        }

        private static bool IsOption(string arg) => arg.StartsWith("-", StringComparison.Ordinal) && arg.Length > 1;

        private static bool TryTakeValue(string[] args, ref int index, out string? value, out string? error)
        {
            // 値の代わりに次のオプションが書かれている誤り(--document-name --allow-font-fallback など)を、値として取り込まない。
            if (index + 1 >= args.Length || IsOption(args[index + 1]))
            {
                value = null;
                error = $"{args[index]} には値が必要です。";
                return false;
            }

            value = args[++index];
            error = null;
            return true;
        }

        private static bool TryTakeNonEmptyValue(string[] args, ref int index, out string? value, out string? error)
        {
            var option = args[index];
            if (!TryTakeValue(args, ref index, out value, out error))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                error = $"{option} の値が空です。";
                return false;
            }

            return true;
        }

        /// <summary>
        /// <c>&lt;セル&gt;=&lt;値&gt;</c> を分割する。セル番地は A1 形式で解釈し、<c>--text</c>・<c>--number</c> を合わせて
        /// 同じセル(表記が違っても同じセル)の重複を使い方エラーにする(要件5.4)。
        /// </summary>
        private static bool TryParseCellValue(
            string pair,
            string option,
            string formatHint,
            Dictionary<CellAddress, string> texts,
            Dictionary<CellAddress, double> numbers,
            out CellAddress cell,
            out string value,
            out string? error)
        {
            cell = default;
            value = string.Empty;
            var separator = pair.IndexOf('=');
            if (separator <= 0)
            {
                // 値(宛名などの個人情報になりうる)を使い方エラーにそのまま出さない。
                error = $"{option} は {formatHint} の形式で指定してください。";
                return false;
            }

            var address = pair.Substring(0, separator);
            if (!CellAddress.TryParse(address, out cell))
            {
                error = $"{option} のセル番地を A1 形式として解釈できません: '{DiagnosticText.Normalize(address, 16)}'";
                return false;
            }

            if (texts.ContainsKey(cell) || numbers.ContainsKey(cell))
            {
                error = $"セル {cell} が重複して指定されています。1つにしてください。";
                return false;
            }

            value = pair.Substring(separator + 1);
            error = null;
            return true;
        }
    }
}
