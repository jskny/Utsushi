using System;
using System.IO;

namespace Utsushi.Parsing.Tests;

/// <summary>テストからリポジトリ内のサンプルを参照するためのパス解決。</summary>
internal static class TestPaths
{
    public static string RepositoryRoot { get; } = FindRepositoryRoot();

    public static string SampleReportsRoot { get; } = Path.Combine(RepositoryRoot, "samples", "reports");

    /// <summary>帳票サンプルのテンプレート(.xlsx)のパスを返す。</summary>
    public static string SampleTemplate(string reportCode) =>
        Path.Combine(SampleReportsRoot, reportCode, "template.xlsx");

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Utsushi.sln"))
                || File.Exists(Path.Combine(directory.FullName, "Utsushi.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException(
            $"リポジトリのルート(Utsushi.sln)が見つかりません。探索開始: {AppContext.BaseDirectory}");
    }
}
