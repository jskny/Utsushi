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

/// <summary>
/// セル番地を直接指定した上書き(帳票定義の置換キーを経由しない経路)で、
/// 指定された文字列がA1形式のセル番地として解釈できない。要件2.8。
/// </summary>
public sealed class InvalidCellOverrideAddressException : UtsushiException
{
    public InvalidCellOverrideAddressException(
        string address,
        string message,
        string? reportCode = null,
        string? sheetName = null)
        : base(message, ProcessingStage.Substitution, reportCode, sheetName)
    {
        Address = address;
    }

    /// <summary>A1形式として解釈できなかった、指定されたセル番地の文字列。</summary>
    public string Address { get; }
}

/// <summary>
/// セル番地を直接指定した上書きの対象が、結合セル範囲内の先頭(アンカー)セル以外だった。要件2.9。
/// </summary>
/// <remarks>
/// 描画(Layoutレイヤー)は結合範囲のアンカーセルの値のみを表示するため、アンカー以外への
/// 上書きは値がモデルには反映されてもPDFには一切出力されない「静かなデータ欠落」になる。
/// これを避けるため、Substitutionレイヤーの時点でエラーとする。
/// </remarks>
public sealed class NonAnchorMergedCellOverrideException : UtsushiException
{
    public NonAnchorMergedCellOverrideException(
        CellAddress address,
        CellAddress anchorAddress,
        string message,
        string? reportCode = null,
        string? sheetName = null)
        : base(message, ProcessingStage.Substitution, reportCode, sheetName, address)
    {
        AnchorAddress = anchorAddress;
    }

    /// <summary>指定されたセルが属する結合範囲の先頭(アンカー)セル。上書きするならこちらを指定する。</summary>
    public CellAddress AnchorAddress { get; }
}
