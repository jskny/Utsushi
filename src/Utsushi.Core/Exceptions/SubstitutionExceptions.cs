namespace Utsushi.Core.Exceptions
{
    /// <summary>
    /// 置換値そのものが不正(<c>null</c>、改行以外の制御文字、対になっていないサロゲートを含む)。要件2.10。
    /// </summary>
    /// <remarks>
    /// 制御文字はフォントに字形が無く、PDF上で豆腐(□)や空白として描画される。
    /// 呼び出し元のデータ不備を黙って描画せず、どの値が原因かを返すためにエラーとする。
    /// </remarks>
    public sealed class InvalidSubstitutionValueException : UtsushiException
    {
        public InvalidSubstitutionValueException(
            string target,
            string message,
            string? reportCode = null,
            string? sheetName = null,
            CellAddress? cellAddress = null)
            : base(message, ProcessingStage.Substitution, reportCode, sheetName, cellAddress)
        {
            Target = target;
        }

        /// <summary>不正な値が渡された置換キー、またはセル番地直接指定のセル番地の文字列。</summary>
        public string Target { get; }
    }

    /// <summary>
    /// セル番地を指定した書き込み(<c>SetText</c>・<c>SetValue</c>)で、
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
}
