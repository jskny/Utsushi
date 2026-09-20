using System;
using System.Collections.Generic;
using System.IO;
using Utsushi;
using Utsushi.Core.Exceptions;
using Utsushi.Rendering;

namespace Utsushi.Cli;

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

            using var converter = ReportPdfConverter.CreateDefault(
                options.DefinitionRoot, fontOptions, renderOptions);
            converter.ConvertToFile(
                options.ReportCode, options.InputPath, options.Values, options.OutputPath, options.CellOverrides);

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
        Console.Error.WriteLine("          [--allow-font-fallback [<代替フォント名>]] [--outline-text]");
        Console.Error.WriteLine();
        Console.Error.WriteLine("オプション:");
        Console.Error.WriteLine("  --report, -r        帳票コード(帳票定義のディレクトリ名)");
        Console.Error.WriteLine("  --input, -i         入力するExcelテンプレート(.xlsx)のパス");
        Console.Error.WriteLine("  --output, -o        出力するPDFのパス");
        Console.Error.WriteLine("  --definitions, -d   帳票定義のルートディレクトリ(既定: ./reports)");
        Console.Error.WriteLine("  --set, -s           置換キーと値。複数指定可(例: --set InvoiceNo=A-001)");
        Console.Error.WriteLine("  --override          セル番地と値。帳票定義への登録有無に関わらず直接上書きする。");
        Console.Error.WriteLine("                      複数指定可(例: --override B5=INV-0001)。結合セルは");
        Console.Error.WriteLine("                      先頭(アンカー)セルの番地を指定すること");
        Console.Error.WriteLine("  --allow-font-fallback");
        Console.Error.WriteLine("                      フォント未検出時に代替フォントを使う(見た目が崩れる可能性あり)");
        Console.Error.WriteLine("  --outline-text      文字をアウトライン化して出力する。ファイルサイズは大幅に小さくなるが");
        Console.Error.WriteLine("                      PDF内の文字列検索・コピーができなくなる");
        Console.Error.WriteLine("  --help, -h          このヘルプを表示する");
    }

    /// <summary>コマンドライン引数。</summary>
    private sealed class CommandLineOptions
    {
        private CommandLineOptions(
            string reportCode,
            string inputPath,
            string outputPath,
            string definitionRoot,
            IReadOnlyDictionary<string, string> values,
            IReadOnlyDictionary<string, string> cellOverrides,
            bool allowFontFallback,
            string? fallbackFont,
            bool outlineText)
        {
            ReportCode = reportCode;
            InputPath = inputPath;
            OutputPath = outputPath;
            DefinitionRoot = definitionRoot;
            Values = values;
            CellOverrides = cellOverrides;
            AllowFontFallback = allowFontFallback;
            FallbackFont = fallbackFont;
            OutlineText = outlineText;
        }

        public string ReportCode { get; }

        public string InputPath { get; }

        public string OutputPath { get; }

        public string DefinitionRoot { get; }

        public IReadOnlyDictionary<string, string> Values { get; }

        public IReadOnlyDictionary<string, string> CellOverrides { get; }

        public bool AllowFontFallback { get; }

        public string? FallbackFont { get; }

        public bool OutlineText { get; }

        /// <summary>引数を解釈する。ヘルプ要求時は options=null, error=null を返す。</summary>
        public static CommandLineOptions? Parse(string[] args, out string? error)
        {
            error = null;

            string? reportCode = null, inputPath = null, outputPath = null;
            var definitionRoot = Path.Combine(Directory.GetCurrentDirectory(), "reports");
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            var cellOverrides = new Dictionary<string, string>(StringComparer.Ordinal);
            var allowFallback = false;
            string? fallbackFont = null;
            var outlineText = false;

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
                        break;

                    case "--set" or "-s":
                        {
                            if (!TryTakeValue(args, ref i, "--set", out var pair, out error)) { return null; }

                            var separator = pair!.IndexOf('=');
                            if (separator <= 0)
                            {
                                error = $"--set の指定は <キー>=<値> の形式で指定してください: '{pair}'";
                                return null;
                            }

                            values[pair.Substring(0, separator)] = pair.Substring(separator + 1);
                            break;
                        }

                    case "--override":
                        {
                            if (!TryTakeValue(args, ref i, "--override", out var pair, out error)) { return null; }

                            var separator = pair!.IndexOf('=');
                            if (separator <= 0)
                            {
                                error = $"--override の指定は <セル番地>=<値> の形式で指定してください: '{pair}'";
                                return null;
                            }

                            cellOverrides[pair.Substring(0, separator)] = pair.Substring(separator + 1);
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

                    case "--outline-text":
                        outlineText = true;
                        break;

                    default:
                        error = $"不明なオプションです: '{args[i]}'";
                        return null;
                }
            }

            if (string.IsNullOrWhiteSpace(reportCode))
            {
                error = "--report は必須です。";
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
                reportCode!, inputPath!, outputPath!, definitionRoot, values, cellOverrides,
                allowFallback, fallbackFont, outlineText);
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
    }
}
