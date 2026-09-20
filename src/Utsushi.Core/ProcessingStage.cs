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
        Parsing,

        /// <summary>帳票定義のロードと <c>WorkbookModel</c> との突合。</summary>
        ReportDefinition,

        /// <summary>置換キーに基づくセル値の差し込み。</summary>
        Substitution,

        /// <summary>ページ分割・座標計算。</summary>
        Layout,

        /// <summary>PDF描画・出力。</summary>
        Rendering,
    }
}
