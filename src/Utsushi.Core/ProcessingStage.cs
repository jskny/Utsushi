namespace Utsushi.Core
{
    /// <summary>
    /// 変換パイプラインの処理段階。例外にどの段階で失敗したかを持たせるために使う(要件6.4)。
    /// </summary>
    public enum ProcessingStage
    {
        /// <summary>段階不明・パイプライン外。</summary>
        Unknown = 0,

        /// <summary>Excel(OOXML)の読み取り。</summary>
        Parsing = 1,

        // 2 は帳票定義の段階(ReportDefinition)だった。帳票定義の廃止(excel2pdf-consolidation 要件1.3)で削除した。
        // ログなどに整数で残った値の意味を変えないよう、後の段階の値は変えない。

        /// <summary>セル番地を指定した値の書き込み(<c>SetText</c>・<c>SetValue</c>)。</summary>
        Substitution = 3,

        /// <summary>ページ分割・座標計算。</summary>
        Layout = 4,

        /// <summary>PDF描画・出力。</summary>
        Rendering = 5,
    }
}
