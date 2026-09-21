using System;
using System.IO;
using System.Text;
using Utsushi.TestSupport;
using Xunit.Sdk;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// ゴールデンファイル(レビュー済みの期待出力)の読み書きと比較。
    /// </summary>
    /// <remarks>
    /// 環境変数 <c>UTSUSHI_UPDATE_GOLDEN=1</c> を付けて実行すると期待値を更新する。
    /// 更新した差分は必ずレビューしてからコミットすること(要件8.3)。
    /// </remarks>
    internal static class GoldenFile
    {
        private const string UpdateEnvironmentVariable = "UTSUSHI_UPDATE_GOLDEN";

        private static bool ShouldUpdate =>
            Environment.GetEnvironmentVariable(UpdateEnvironmentVariable) is "1" or "true";

        /// <summary>
        /// 実際の出力を、リポジトリ内のゴールデンファイルと比較する。
        /// </summary>
        /// <param name="reportCode">帳票コード(ゴールデンファイルの配置ディレクトリ)。</param>
        /// <param name="fileName">ゴールデンファイル名。</param>
        /// <param name="actual">実際の出力。</param>
        public static void Verify(string reportCode, string fileName, string actual)
        {
            var path = Path.Combine(TestPaths.FixturesRoot, reportCode, fileName);
            var normalized = Normalize(actual);

            if (ShouldUpdate || !File.Exists(path))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                File.WriteAllText(path, normalized, new UTF8Encoding(false));

                if (!ShouldUpdate)
                {
                    // 期待値が無いまま「一致した」と報告しないよう、初回は必ず失敗させる。
                    throw new XunitException(
                        $"ゴールデンファイルが存在しなかったため新規作成しました: {path}{Environment.NewLine}"
                        + "内容をレビューしてコミットしたうえで、テストを再実行してください。");
                }

                return;
            }

            var expected = Normalize(File.ReadAllText(path));
            if (string.Equals(expected, normalized, StringComparison.Ordinal))
            {
                return;
            }

            // 失敗時に「どこが最初に食い違ったか」を出す。全文diffはレビューしづらいため。
            throw new XunitException(BuildFailureMessage(path, expected, normalized));
        }

        /// <summary>改行コードの差と末尾の空白行を吸収する。</summary>
        private static string Normalize(string text) =>
            text.Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd('\n') + "\n";

        private static string BuildFailureMessage(string path, string expected, string actual)
        {
            var expectedLines = expected.Split('\n');
            var actualLines = actual.Split('\n');

            var sb = new StringBuilder();
            sb.AppendLine($"ゴールデンファイルと一致しません: {path}");
            sb.AppendLine($"期待 {expectedLines.Length} 行 / 実際 {actualLines.Length} 行");

            var max = Math.Max(expectedLines.Length, actualLines.Length);
            var shown = 0;

            for (var i = 0; i < max && shown < 10; i++)
            {
                var e = i < expectedLines.Length ? expectedLines[i] : "(行なし)";
                var a = i < actualLines.Length ? actualLines[i] : "(行なし)";
                if (string.Equals(e, a, StringComparison.Ordinal))
                {
                    continue;
                }

                sb.AppendLine($"  L{i + 1} 期待: {e}");
                sb.AppendLine($"  L{i + 1} 実際: {a}");
                shown++;
            }

            sb.AppendLine();
            sb.AppendLine(
                $"意図した変更であれば、{UpdateEnvironmentVariable}=1 を付けて実行し、差分をレビューしてコミットしてください。");
            return sb.ToString();
        }
    }
}
