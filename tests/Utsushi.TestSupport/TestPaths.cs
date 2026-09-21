using System;
using System.IO;

namespace Utsushi.TestSupport
{
    /// <summary>各テストプロジェクトからリポジトリ内のサンプル・ゴールデンファイルを参照するためのパス解決。</summary>
    public static class TestPaths
    {
        public static string RepositoryRoot { get; } = FindRepositoryRoot();

        public static string SampleReportsRoot { get; } = Path.Combine(RepositoryRoot, "samples", "reports");

        /// <summary>ゴールデンファイルの配置ルート(`.kiro/steering/structure.md`「命名規則」)。</summary>
        public static string FixturesRoot { get; } =
            Path.Combine(RepositoryRoot, "tests", "Utsushi.Golden.Tests", "Fixtures");

        /// <summary>帳票サンプルのテンプレート(.xlsx)のパスを返す。</summary>
        public static string SampleTemplate(string reportCode) =>
            Path.Combine(SampleReportsRoot, reportCode, "template.xlsx");

        /// <summary>
        /// テスト実行ディレクトリ(bin/Debug/netX.0)から上へ辿り、リポジトリのルートを探す。
        /// </summary>
        /// <remarks>
        /// `Utsushi.sln`(classic形式)はVS2019対応のため `vs2019/` 配下に置かれ、ルート直下には
        /// `Utsushi.slnx` のみが存在する(`.kiro/steering/structure.md`「ソリューション構成」参照)。
        /// </remarks>
        private static string FindRepositoryRoot()
        {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory is not null)
            {
                if (File.Exists(Path.Combine(directory.FullName, "Utsushi.slnx")))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            throw new InvalidOperationException(
                $"リポジトリのルート(Utsushi.slnx)が見つかりません。探索開始: {AppContext.BaseDirectory}");
        }
    }
}
