using Utsushi.Layout.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// <see cref="ReportModel"/> から、ページ分割済み・座標確定済みの <see cref="PagedLayout"/> を計算する。
    /// </summary>
    public interface IReportLayoutEngine
    {
        /// <summary>レイアウトを計算する。文字の収まりの確認(要件13)も行い、<see cref="PagedLayout.FitIssues"/> に入れる。</summary>
        PagedLayout Compute(ReportModel report);

        /// <summary>
        /// レイアウトを計算する。<paramref name="checkFit"/> が false なら文字の収まりの確認(要件13)を省く
        /// (PDFへの変換では結果を使わないため、文字幅の計測を減らす)。
        /// </summary>
        /// <remarks>既定の実装は <see cref="Compute(ReportModel)"/> を呼ぶ(既存の実装を壊さないため)。</remarks>
        PagedLayout Compute(ReportModel report, bool checkFit) => Compute(report);
    }
}
