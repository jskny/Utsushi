using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Core.Exceptions;

namespace Utsushi.Cli
{
    /// <summary>
    /// Utsushi のコマンドラインエントリーポイント(excel2pdf-consolidation 要件5)。
    /// </summary>
    /// <remarks>
    /// 本体は <see cref="Excel2Pdf"/>(ライブラリ)であり、CLI はテンプレートの診断と変換を試すための薄い入口である。
    /// 呼び出し元プロダクトは CLI ではなく <c>Utsushi</c> アセンブリを直接参照する。
    /// </remarks>
    internal static class Program
    {
        /// <summary><c>check</c>: 問題なし。<c>convert</c>: 保存した。<c>--help</c>: 使い方を出した。</summary>
        internal const int ExitSuccess = 0;

        /// <summary><c>check</c>: 問題が1つ以上ある。</summary>
        internal const int ExitIssuesFound = 1;

        /// <summary>使い方エラー。</summary>
        internal const int ExitUsageError = 2;

        /// <summary>変換・診断のエラー。</summary>
        internal const int ExitConversionError = 3;

        /// <summary>エラー出力の「原因:」の行の最大の長さ。</summary>
        private const int MaxCauseLength = 200;

        internal static int Main(string[] args) => Run(args, Console.Out, Console.Error);

        /// <summary>CLI を実行し、終了コードを返す(design.md 5.3)。</summary>
        internal static int Run(string[] args, TextWriter stdout, TextWriter stderr)
        {
            var options = CommandLineOptions.Parse(args, out var error);
            if (options is null)
            {
                if (error is null)
                {
                    PrintUsage(stdout);
                    return ExitSuccess;
                }

                stderr.WriteLine("エラー: " + error);
                stderr.WriteLine();
                PrintUsage(stderr);
                return ExitUsageError;
            }

            var action = options.Command == CliCommand.Check ? "診断" : "変換";
            try
            {
                using var pdf = new Excel2Pdf(options.InputPath, options.ToExcel2PdfOptions());
                foreach (var (cell, text) in options.Texts)
                {
                    pdf.SetText(cell.ToString(), text);
                }

                foreach (var (cell, number) in options.Numbers)
                {
                    pdf.SetValue(cell.ToString(), number);
                }

                if (options.Command == CliCommand.Check)
                {
                    var diagnostics = pdf.Check();
                    WriteDiagnostics(diagnostics, stdout);
                    return CheckExitCode(diagnostics);
                }

                pdf.Save(options.OutputPath!);
                stdout.WriteLine($"PDFを出力しました: {options.OutputPath}");
                return ExitSuccess;
            }
            catch (UtsushiException ex)
            {
                // 文書名・シート名・セル番地・処理段階を含めて報告する。未対応要素による中止では要素の一覧を含む(要件5.7)。
                stderr.WriteLine($"{action}に失敗しました。");
                stderr.WriteLine(ex.ToDiagnosticString());
                if (ex.InnerException is not null)
                {
                    // 内側の例外のメッセージは、ファイル由来の不正な値を含みうる。改行で偽の行を差し込ませず、長さを抑える(security-reviewer指摘)。
                    stderr.WriteLine("原因: " + DiagnosticText.Normalize(ex.InnerException.Message, MaxCauseLength));
                }

                return ExitConversionError;
            }
            catch (InvalidOperationException ex)
            {
                // 部品の契約違反など、入力によらない内部のエラー。プロセスを異常終了させず、変換・診断のエラーとして報告する。
                stderr.WriteLine($"{action}に失敗しました(内部のエラー): " + DiagnosticText.Normalize(ex.Message, MaxCauseLength));
                return ExitConversionError;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                stderr.WriteLine("ファイルの入出力に失敗しました: " + ex.Message);
                return ExitConversionError;
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
            {
                // 引数の解釈では弾けない不正な値(形式が不正なパスなど)を、未処理の例外でプロセスを落とさず使い方エラーとして報告する。
                stderr.WriteLine("エラー: 引数の値が不正です: " + ex.Message);
                stderr.WriteLine("使い方は --help を参照してください。");
                return ExitUsageError;
            }
        }

        /// <summary>
        /// <c>check</c> の終了コード。問題(未対応要素・文字の収まらない箇所・フォントの代替)が1つ以上あれば1、無ければ0。
        /// 書式の近似は終了コードに影響させない(要件5.2)。
        /// </summary>
        internal static int CheckExitCode(Excel2PdfDiagnostics diagnostics) =>
            diagnostics.HasIssues ? ExitIssuesFound : ExitSuccess;

        /// <summary>
        /// 診断の結果を種類ごとに出す(design.md 5.2)。書式の近似は「(参考)」として最後に出す。
        /// 文字の収まらない箇所は、セルの文字列(個人情報を含みうる)を出さず、説明の文だけを出す。
        /// </summary>
        internal static void WriteDiagnostics(Excel2PdfDiagnostics diagnostics, TextWriter output)
        {
            WriteSection(output, "未対応の要素", diagnostics.UnsupportedElements, diagnostics.UnsupportedElementsTruncated, e => e.ToString());
            WriteSection(
                output,
                "文字の収まらない箇所",
                diagnostics.FitIssues,
                diagnostics.FitIssuesTruncated,
                i => $"[{i.Kind}] {i.Cell}({i.PageNumber}ページ目): {DiagnosticText.Normalize(i.Message, int.MaxValue)}");
            WriteSection(output, "フォントの代替", diagnostics.FontSubstitutions, truncated: false, s => s.ToString());
            WriteSection(
                output,
                "書式の近似(参考)",
                diagnostics.FormatApproximations,
                diagnostics.FormatApproximationsTruncated,
                a => a.ToString());

            if (!diagnostics.HasIssues)
            {
                output.WriteLine(diagnostics.HasFormatApproximations
                    ? "問題は見つかりませんでした(書式の近似は参考情報です)。"
                    : "問題は見つかりませんでした。");
            }
        }

        private static void WriteSection<T>(
            TextWriter output, string title, IReadOnlyList<T> items, bool truncated, Func<T, string> format)
        {
            if (items.Count == 0)
            {
                return;
            }

            output.WriteLine($"{title}: {items.Count}件" + (truncated ? "(上限で打ち切り)" : string.Empty));
            foreach (var line in items.Select(format))
            {
                output.WriteLine("  " + line);
            }
        }

        private static void PrintUsage(TextWriter output)
        {
            output.WriteLine("Utsushi - Excel のテンプレートの診断と PDF への変換");
            output.WriteLine();
            output.WriteLine("使い方:");
            output.WriteLine("  utsushi check   <入力.xlsx>              [オプション]");
            output.WriteLine("  utsushi convert <入力.xlsx> <出力.pdf>   [オプション]");
            output.WriteLine("  utsushi --help");
            output.WriteLine();
            output.WriteLine("  check   未対応の要素・文字の収まらない箇所・フォントの代替・書式の近似(参考)を出す。");
            output.WriteLine("          問題が無ければ終了コード0、あれば1(書式の近似だけなら0)");
            output.WriteLine("  convert PDF に変換して保存する");
            output.WriteLine("  パスはコマンドの直後(オプションより前)に書く");
            output.WriteLine();
            output.WriteLine("オプション:");
            output.WriteLine("  --text <セル>=<文字列>          セルに文字列を書き込む(複数指定可。例: --text A3=株式会社サンプル)");
            output.WriteLine("  --number <セル>=<数値>          セルに数値を書き込む。セルの表示形式で表示する(複数指定可。");
            output.WriteLine("                                  小数点は「.」。日時は Excel のシリアル値で指定する)");
            output.WriteLine("  --sheet <シート名>              変換するシート(省略すると保存時に開いていたシート)");
            output.WriteLine("  --document-name <名前>          PDF のタイトル・エラーの情報に使う名前(省略すると入力のファイル名)");
            output.WriteLine("  --max-digit-width <px>          列幅の換算に使う最大数字幅(省略するとブックの標準フォントから見積もる)");
            output.WriteLine("  --ignore-unsupported            グラフなどの未対応の要素があっても、その要素を除いて変換する");
            output.WriteLine("                                  (check では、指定しても問題として数える)");
            output.WriteLine("  --allow-font-fallback [書体名]  フォントが無ければ代替の書体で描く(省略すると同梱のBIZ UDPゴシック)");
            output.WriteLine("  --allow-missing-glyphs          フォントに字形の無い文字があってもエラーにしない(動作確認用)");
            output.WriteLine("  --outline-text                  文字をアウトライン化する(PDF の文字列の検索・コピーはできなくなる)");
            output.WriteLine("  -h, --help                      この使い方を出す");
            output.WriteLine();
            output.WriteLine("終了コード: 0 成功/問題なし、1 問題あり(check)、2 使い方エラー、3 変換・診断のエラー");
        }
    }
}
