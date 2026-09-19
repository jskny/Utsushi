using System;
using System.Collections.Generic;
using System.IO;
using Utsushi.Core.Exceptions;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.ReportDefinitions;

/// <summary>
/// 帳票コードから帳票定義を取得する。
/// </summary>
public interface IReportDefinitionRepository
{
    /// <summary>帳票コードに対応する帳票定義をロードする。</summary>
    /// <exception cref="ReportDefinitionNotFoundException">定義が存在しない場合(要件1.4)。</exception>
    /// <exception cref="ReportDefinitionSchemaException">定義が不正な場合(要件6.3)。</exception>
    ReportDefinition Load(string reportCode);

    /// <summary>登録済みの帳票コードを列挙する。</summary>
    IReadOnlyCollection<string> ListReportCodes();
}

/// <summary>
/// <c>&lt;ルート&gt;/&lt;帳票コード&gt;/definition.json</c> の配置規約に従って帳票定義を読み込むリポジトリ。
/// </summary>
/// <remarks>
/// 配置規約は `.kiro/steering/structure.md`「命名規則」に従う。
/// 新しい帳票への対応は、このディレクトリへの定義追加のみで完結する(要件8.1)。
/// </remarks>
public sealed class FileSystemReportDefinitionRepository : IReportDefinitionRepository
{
    /// <summary>帳票定義ファイルの固定ファイル名。</summary>
    public const string DefinitionFileName = "definition.json";

    private readonly string _rootDirectory;

    public FileSystemReportDefinitionRepository(string rootDirectory)
    {
        if (string.IsNullOrWhiteSpace(rootDirectory))
        {
            throw new ArgumentException("帳票定義のルートディレクトリが空です。", nameof(rootDirectory));
        }

        _rootDirectory = Path.GetFullPath(rootDirectory);
    }

    /// <inheritdoc />
    public ReportDefinition Load(string reportCode)
    {
        var path = ResolveDefinitionPath(reportCode);

        if (!File.Exists(path))
        {
            throw new ReportDefinitionNotFoundException(
                reportCode,
                $"帳票コード '{reportCode}' の帳票定義が見つかりません(期待パス: {path})。");
        }

        string json;
        try
        {
            json = File.ReadAllText(path);
        }
        catch (IOException ex)
        {
            throw new ReportDefinitionNotFoundException(
                reportCode, $"帳票定義ファイルを読み込めません: {path}", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new ReportDefinitionNotFoundException(
                reportCode, $"帳票定義ファイルへのアクセスが拒否されました: {path}", ex);
        }

        var definition = ReportDefinitionJsonReader.Read(json, path);

        if (!string.Equals(definition.ReportCode, reportCode, StringComparison.Ordinal))
        {
            throw new ReportDefinitionSchemaException(
                $"帳票定義の reportCode ('{definition.ReportCode}') が、"
                + $"配置ディレクトリ名 ('{reportCode}') と一致しません。",
                reportCode,
                path,
                "reportCode");
        }

        return definition;
    }

    /// <inheritdoc />
    public IReadOnlyCollection<string> ListReportCodes()
    {
        if (!Directory.Exists(_rootDirectory))
        {
            return Array.Empty<string>();
        }

        var codes = new List<string>();
        foreach (var directory in Directory.EnumerateDirectories(_rootDirectory))
        {
            if (File.Exists(Path.Combine(directory, DefinitionFileName)))
            {
                codes.Add(Path.GetFileName(directory));
            }
        }

        codes.Sort(StringComparer.Ordinal);
        return codes;
    }

    /// <summary>
    /// 帳票コードから定義ファイルのパスを解決する。
    /// </summary>
    /// <remarks>
    /// 帳票コードは外部(呼び出し元プロダクト)から渡されうるため、
    /// パス区切り文字や <c>..</c> を含む値でルートディレクトリの外へ出られないことを検証する。
    /// </remarks>
    private string ResolveDefinitionPath(string reportCode)
    {
        if (string.IsNullOrWhiteSpace(reportCode))
        {
            throw new ReportDefinitionNotFoundException(
                reportCode ?? string.Empty, "帳票コードが空です。");
        }

        if (reportCode.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
            || reportCode.IndexOf('/') >= 0
            || reportCode.IndexOf('\\') >= 0
            || reportCode == "."
            || reportCode == "..")
        {
            throw new ReportDefinitionNotFoundException(
                reportCode,
                $"帳票コード '{reportCode}' にディレクトリ名として使用できない文字が含まれています。");
        }

        var candidate = Path.GetFullPath(Path.Combine(_rootDirectory, reportCode, DefinitionFileName));
        var expectedPrefix = _rootDirectory.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? _rootDirectory
            : _rootDirectory + Path.DirectorySeparatorChar;

        if (!candidate.StartsWith(expectedPrefix, StringComparison.Ordinal))
        {
            throw new ReportDefinitionNotFoundException(
                reportCode,
                $"帳票コード '{reportCode}' は帳票定義ルートの外を指しています。");
        }

        return candidate;
    }
}
