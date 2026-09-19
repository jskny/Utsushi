using System;
using System.IO;

namespace Utsushi.ReportDefinition.Tests;

/// <summary>テストからリポジトリ内のサンプルを参照するためのパス解決。</summary>
internal static class TestPaths
{
    /// <summary>リポジトリのルートディレクトリ。</summary>
    public static string RepositoryRoot { get; } = FindRepositoryRoot();

    /// <summary>帳票サンプルのルート(<c>samples/reports</c>)。</summary>
    public static string SampleReportsRoot { get; } = Path.Combine(RepositoryRoot, "samples", "reports");

    /// <summary>
    /// テスト実行ディレクトリ(bin/Debug/netX.0)から上へ辿り、リポジトリのルートを探す。
    /// </summary>
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
