using System;
using System.Collections.Generic;
using System.IO;
using Utsushi;
using Utsushi.Core.Exceptions;
using Utsushi.Rendering;

namespace Utsushi.Cli
{
    /// <summary>
    /// Utsushi のコマンドラインエントリーポイント。
    /// </summary>
    /// <remarks>
    /// 本体は <see cref="ReportPdfConverter"/>(ライブラリ)であり、CLI は動作確認・バッチ実行用の薄い入口である。
    /// 呼び出し元プロダクトは CLI ではなく <c>Utsushi</c> アセンブリを直接参照する。
    /// </remarks>
    internal static class Program
    {
        private const int ExitSuccess = 0;
        private const int ExitUsageError = 2;
        private const int ExitConversionError = 3;

        private static int Main(string[] args)
        {
            var options = CommandLineOptions.Parse(args, out var error);
            if (options is null)
            {
                if (error is not null)
                {
                    Console.Error.WriteLine("エラー: " + error);
                    Console.Error.WriteLine();
                }

                PrintUsage();
                return error is null ? ExitSuccess : ExitUsageError;
            }

            try
            {
                var fontOptions = options.AllowFontFallback
                    ? FontResolverOptions.AllowFallback(options.FallbackFont)
                    : FontResolverOptions.Strict;

                var renderOptions = options.OutlineText ? PdfRenderOptions.OutlineText : PdfRenderOptions.Default;
                if (options.AllowMissingGlyphs)
                {
                    renderOptions = renderOptions with { MissingGlyphs = MissingGlyphPolicy.Render };
                }

                if (options.ReportCode is null)
                {
                    // --report 省略時は帳票定義なしで変換する(要件12.7)。
                    using var converter = ReportPdfConverter.CreateDefault(fontOptions, renderOptions);
                    converter.ConvertFileWithoutDefinition(
                        options.InputPath, options.OutputPath, options.CellOverrides, maxDigitWidthPx: options.MaxDigitWidthPx);
                }
                else
                {
                    using var converter = ReportPdfConverter.CreateDefault(
                        options.DefinitionRoot, fontOptions, renderOptions);
                    converter.ConvertToFile(
                        options.ReportCode, options.InputPath, options.Values, options.OutputPath, options.CellOverrides);
                }

                Console.WriteLine($"PDFを出力しました: {options.OutputPath}");
                return ExitSuccess;
            }
            catch (UtsushiException ex)
            {
                // 帳票コード・シート名・セル番地・処理段階を含めて報告する(要件6.4)。
                Console.Error.WriteLine("変換に失敗しました。");
                Console.Error.WriteLine(ex.ToDiagnosticString());
                if (ex.InnerException is not null)
                {
                    Console.Error.WriteLine("原因: " + ex.InnerException.Message);
                }

                return ExitConversionError;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Console.Error.WriteLine("ファイル入出力に失敗しました: " + ex.Message);
                return ExitConversionError;
            }
        }

        private static void PrintUsage()
        {
            Console.Error.WriteLine("Utsushi - 自社帳票(Excel)のPDF変換");
            Console.Error.WriteLine();
            Console.Error.WriteLine("使い方:");
            Console.Error.WriteLine("  utsushi --report <帳票コード> --input <xlsxパス> --output <pdfパス>");
            Console.Error.WriteLine("          [--definitions <帳票定義ルート>] [--set <キー>=<値> ...]");
            Console.Error.WriteLine("          [--override <セル番地>=<値> ...]");
            Console.Error.WriteLine("          [--allow-font-fallback [<代替フォント名>]] [--allow-missing-glyphs] [--outline-text]");
            Console.Error.WriteLine("  utsushi --input <xlsxパス> --output <pdfパス> [--override <セル番地>=<値> ...] [上記の出力オプション]");
            Console.Error.WriteLine("          (帳票定義なしで変換する。アクティブシートを変換し、未対応の要素は無視する)");
            Console.Error.WriteLine();
            Console.Error.WriteLine("オプション:");
            Console.Error.WriteLine("  --report, -r        帳票コード(帳票定義のディレクトリ名)。省略すると帳票定義なしで変換する");
            Console.Error.WriteLine("                      (差し込みの --set は使えない。見た目の一致は保証しない)");
            Console.Error.WriteLine("  --input, -i         入力するExcelテンプレート(.xlsx)のパス");
            Console.Error.WriteLine("  --output, -o        出力するPDFのパス");
            Console.Error.WriteLine("  --definitions, -d   帳票定義のルートディレクトリ(既定: ./reports)");
            Console.Error.WriteLine("  --set, -s           置換キーと値。複数指定可(例: --set InvoiceNo=A-001)");
            Console.Error.WriteLine("  --override          セル番地と値。帳票定義への登録有無に関わらず直接上書きする。");
            Console.Error.WriteLine("                      複数指定可(例: --override A1=請求書(控))。結合セルは");
            Console.Error.WriteLine("                      先頭(アンカー)セルの番地を指定すること");
            Console.Error.WriteLine("  --max-digit-width   帳票定義なしの変換で、列幅の換算に使う最大数字幅(ピクセル)。");
            Console.Error.WriteLine("                      省略時はブックの標準フォントから見積もる(Calibri 11=7、");
            Console.Error.WriteLine("                      ＭＳ Ｐゴシック/游ゴシック 11=8)。列幅がExcelとずれる場合に指定する");
            Console.Error.WriteLine("  --allow-font-fallback");
            Console.Error.WriteLine("                      フォント未検出時に代替フォントを使う(見た目が崩れる可能性あり)");
            Console.Error.WriteLine("  --allow-missing-glyphs");
            Console.Error.WriteLine("                      フォントに字形が無い文字があってもエラーにせず出力する");
            Console.Error.WriteLine("                      (その文字は豆腐(□)や空白になる。動作確認用)");
            Console.Error.WriteLine("  --outline-text      文字をアウトライン化して出力する。ファイルサイズは大幅に小さくなるが");
            Console.Error.WriteLine("                      PDF内の文字列検索・コピーができなくなる");
            Console.Error.WriteLine("  --help, -h          このヘルプを表示する");
        }

        /// <summary>コマンドライン引数。</summary>
        internal sealed class CommandLineOptions
        {
            private CommandLineOptions(
                string? reportCode,
                string inputPath,
                string outputPath,
                string definitionRoot,
                IReadOnlyDictionary<string, string> values,
                IReadOnlyDictionary<string, string> cellOverrides,
                bool allowFontFallback,
                string? fallbackFont,
                bool allowMissingGlyphs,
                bool outlineText,
                double? maxDigitWidthPx)
            {
                MaxDigitWidthPx = maxDigitWidthPx;
                ReportCode = reportCode;
                InputPath = inputPath;
                OutputPath = outputPath;
                DefinitionRoot = definitionRoot;
                Values = values;
                CellOverrides = cellOverrides;
                AllowFontFallback = allowFontFallback;
                FallbackFont = fallbackFont;
                AllowMissingGlyphs = allowMissingGlyphs;
                OutlineText = outlineText;
            }

            /// <summary>帳票コード。null の場合は帳票定義なしで変換する(要件12.7)。</summary>
            public string? ReportCode { get; }

            public string InputPath { get; }

            public string OutputPath { get; }

            public string DefinitionRoot { get; }

            public IReadOnlyDictionary<string, string> Values { get; }

            public IReadOnlyDictionary<string, string> CellOverrides { get; }

            public bool AllowFontFallback { get; }

            public string? FallbackFont { get; }

            public bool AllowMissingGlyphs { get; }

            public bool OutlineText { get; }

            /// <summary>帳票定義なしの変換で使う最大数字幅。null ならブックの標準フォントから見積もる。</summary>
            public double? MaxDigitWidthPx { get; }

            /// <summary>引数を解釈する。ヘルプ要求時は options=null, error=null を返す。</summary>
            public static CommandLineOptions? Parse(string[] args, out string? error)
            {
                error = null;

                string? reportCode = null, inputPath = null, outputPath = null;
                var definitionRoot = Path.Combine(Directory.GetCurrentDirectory(), "reports");
                var definitionRootSpecified = false;
                var values = new Dictionary<string, string>(StringComparer.Ordinal);
                var cellOverrides = new Dictionary<string, string>(StringComparer.Ordinal);
                var allowFallback = false;
                string? fallbackFont = null;
                var allowMissingGlyphs = false;
                var outlineText = false;
                double? maxDigitWidthPx = null;

                if (args.Length == 0)
                {
                    return null;
                }

                for (var i = 0; i < args.Length; i++)
                {
                    switch (args[i])
                    {
                        case "--help" or "-h":
                            return null;

                        case "--report" or "-r":
                            if (!TryTakeValue(args, ref i, "--report", out reportCode, out error)) { return null; }
                            break;

                        case "--input" or "-i":
                            if (!TryTakeValue(args, ref i, "--input", out inputPath, out error)) { return null; }
                            break;

                        case "--output" or "-o":
                            if (!TryTakeValue(args, ref i, "--output", out outputPath, out error)) { return null; }
                            break;

                        case "--definitions" or "-d":
                            if (!TryTakeValue(args, ref i, "--definitions", out var root, out error)) { return null; }
                            definitionRoot = root!;
                            definitionRootSpecified = true;
                            break;

                        case "--set" or "-s":
                            {
                                if (!TryTakeValue(args, ref i, "--set", out var pair, out error)) { return null; }
                                if (!TryParseKeyValuePair(pair!, "--set", "<キー>=<値>", values, out error)) { return null; }
                                break;
                            }

                        case "--override":
                            {
                                if (!TryTakeValue(args, ref i, "--override", out var pair, out error)) { return null; }
                                if (!TryParseKeyValuePair(pair!, "--override", "<セル番地>=<値>", cellOverrides, out error)) { return null; }
                                break;
                            }

                        case "--allow-font-fallback":
                            allowFallback = true;

                            // 直後の引数がオプションでなければ代替フォント名として扱う。
                            if (i + 1 < args.Length && !args[i + 1].StartsWith("-", StringComparison.Ordinal))
                            {
                                fallbackFont = args[++i];
                            }

                            break;

                        case "--allow-missing-glyphs":
                            allowMissingGlyphs = true;
                            break;

                        case "--outline-text":
                            outlineText = true;
                            break;

                        case "--max-digit-width":
                            {
                                if (!TryTakeValue(args, ref i, "--max-digit-width", out var text, out error)) { return null; }
                                if (!double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var mdw)
                                    || double.IsNaN(mdw) || mdw <= 0 || double.IsInfinity(mdw))
                                {
                                    error = $"--max-digit-width には正の数を指定してください: '{text}'";
                                    return null;
                                }

                                maxDigitWidthPx = mdw;
                                break;
                            }

                        default:
                            error = $"不明なオプションです: '{args[i]}'";
                            return null;
                    }
                }

                if (reportCode is not null && string.IsNullOrWhiteSpace(reportCode))
                {
                    error = "--report の値が空です。帳票定義なしで変換する場合は --report 自体を省略してください。";
                    return null;
                }

                if (reportCode is null && values.Count > 0)
                {
                    error = "--set(置換キーによる差し込み)は --report と併用してください。"
                        + "帳票定義なしで変換する場合は --override <セル番地>=<値> を使ってください。";
                    return null;
                }

                if (reportCode is not null && maxDigitWidthPx is not null)
                {
                    error = "--max-digit-width は帳票定義なしの変換(--report の省略)でのみ指定できます。"
                        + "帳票定義ありの場合は definition.json の maxDigitWidthPx を使います。";
                    return null;
                }

                if (reportCode is null && definitionRootSpecified)
                {
                    error = "--definitions は --report と併用してください。";
                    return null;
                }

                if (string.IsNullOrWhiteSpace(inputPath))
                {
                    error = "--input は必須です。";
                    return null;
                }

                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    error = "--output は必須です。";
                    return null;
                }

                return new CommandLineOptions(
                    reportCode, inputPath!, outputPath!, definitionRoot, values, cellOverrides,
                    allowFallback, fallbackFont, allowMissingGlyphs, outlineText, maxDigitWidthPx);
            }

            private static bool TryTakeValue(
                string[] args, ref int index, string optionName, out string? value, out string? error)
            {
                if (index + 1 >= args.Length)
                {
                    value = null;
                    error = $"{optionName} には値が必要です。";
                    return false;
                }

                value = args[++index];
                error = null;
                return true;
            }

            /// <summary>"<キー>=<値>" 形式の引数を分割し、辞書に格納する(--set / --override で共通)。</summary>
            private static bool TryParseKeyValuePair(
                string pair, string optionName, string formatHint, Dictionary<string, string> destination, out string? error)
            {
                var separator = pair.IndexOf('=');
                if (separator <= 0)
                {
                    error = $"{optionName} の指定は {formatHint} の形式で指定してください: '{pair}'";
                    return false;
                }

                destination[pair.Substring(0, separator)] = pair.Substring(separator + 1);
                error = null;
                return true;
            }
        }
    }
}
