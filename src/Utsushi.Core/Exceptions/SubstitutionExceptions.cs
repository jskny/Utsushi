namespace Utsushi.Core.Exceptions;

/// <summary>
/// 帳票定義に存在しない置換キーが指定された。要件2.3。
/// </summary>
public sealed class SubstitutionKeyNotFoundException : UtsushiException
{
    public SubstitutionKeyNotFoundException(string key, string message, string? reportCode = null, string? sheetName = null)
        : base(message, ProcessingStage.Substitution, reportCode, sheetName)
    {
        Key = key;
    }

    /// <summary>帳票定義に存在しなかった置換キー。</summary>
    public string Key { get; }
}

/// <summary>
/// 帳票定義が必須とした置換キーに対応する値が渡されていない。要件2.4。
/// </summary>
public sealed class RequiredSubstitutionValueMissingException : UtsushiException
{
    public RequiredSubstitutionValueMissingException(
        string key,
        string message,
        string? reportCode = null,
        string? sheetName = null,
        CellAddress? cellAddress = null)
        : base(message, ProcessingStage.Substitution, reportCode, sheetName, cellAddress)
    {
        Key = key;
    }

    /// <summary>値が渡されなかった必須の置換キー。</summary>
    public string Key { get; }
}
