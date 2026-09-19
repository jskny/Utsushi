using Utsushi.Layout.Model;
using Utsushi.ReportDefinitions.Model;

namespace Utsushi.Layout;

/// <summary>
/// <see cref="ReportModel"/> から、ページ分割済み・座標確定済みの <see cref="PagedLayout"/> を計算する。
/// </summary>
public interface IReportLayoutEngine
{
    /// <summary>レイアウトを計算する。</summary>
    PagedLayout Compute(ReportModel report);
}
