using System;

namespace Utsushi.Core.Exceptions;

/// <summary>
/// 指定された帳票コードに対応する帳票定義が見つからない。要件1.4。
/// </summary>
public sealed class ReportDefinitionNotFoundException : UtsushiException
{
    public ReportDefinitionNotFoundException(string reportCode, string message, Exception? innerException = null)
        : base(message, ProcessingStage.ReportDefinition, reportCode, innerException: innerException)
    {
    }
}

/// <summary>
/// 帳票定義ファイル自体が不正(JSONとして壊れている、スキーマ不一致、必須項目欠落など)。要件6.3。
/// </summary>
public sealed class ReportDefinitionSchemaException : UtsushiException
{
    public ReportDefinitionSchemaException(
        string message,
        string? reportCode = null,
        string? definitionPath = null,
        string? propertyPath = null,
        Exception? innerException = null)
        : base(message, ProcessingStage.ReportDefinition, reportCode, innerException: innerException)
    {
        DefinitionPath = definitionPath;
        PropertyPath = propertyPath;
    }

    /// <summary>問題のあった帳票定義ファイルのパス(判明している場合)。</summary>
    public string? DefinitionPath { get; }

    /// <summary>問題のあったJSONプロパティのパス(例: "substitutionFields[0].cell")。</summary>
    public string? PropertyPath { get; }
}

/// <summary>
/// 帳票定義が期待するシート構成・セル番地が、実際のワークブックと一致しない。要件1.4。
/// </summary>
public sealed class ReportStructureMismatchException : UtsushiException
{
    public ReportStructureMismatchException(
        string message,
        string? reportCode = null,
        string? sheetName = null,
        CellAddress? cellAddress = null)
        : base(message, ProcessingStage.ReportDefinition, reportCode, sheetName, cellAddress)
    {
    }
}
