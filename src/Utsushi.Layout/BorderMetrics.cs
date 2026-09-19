using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout;

/// <summary>
/// Excel の罫線スタイルを、線幅(ポイント)と破線パターンへ対応付ける。
/// </summary>
/// <remarks>
/// Excel は罫線を画面上のピクセル数で定義しているため、96dpi 基準でポイントへ換算した値を用いる
/// (hair=1px未満の極細、thin=1px、medium=2px、thick=3px)。
/// </remarks>
internal static class BorderMetrics
{
    /// <summary>二重線の2本の線の間隔(ポイント)。</summary>
    public const double DoubleLineGapPt = 1.0;

    /// <summary>罫線スタイルに対応する線幅(ポイント)。</summary>
    public static double GetWidthPt(BorderLineStyle style) => style switch
    {
        BorderLineStyle.None => 0.0,
        BorderLineStyle.Hair => 0.25,
        BorderLineStyle.Thin => Units.PixelsToPoints(1.0),
        BorderLineStyle.Dotted => Units.PixelsToPoints(1.0),
        BorderLineStyle.Dashed => Units.PixelsToPoints(1.0),
        BorderLineStyle.DashDot => Units.PixelsToPoints(1.0),
        BorderLineStyle.DashDotDot => Units.PixelsToPoints(1.0),
        BorderLineStyle.Medium => Units.PixelsToPoints(2.0),
        BorderLineStyle.MediumDashed => Units.PixelsToPoints(2.0),
        BorderLineStyle.MediumDashDot => Units.PixelsToPoints(2.0),
        BorderLineStyle.MediumDashDotDot => Units.PixelsToPoints(2.0),
        BorderLineStyle.SlantDashDot => Units.PixelsToPoints(2.0),
        BorderLineStyle.Thick => Units.PixelsToPoints(3.0),
        BorderLineStyle.Double => Units.PixelsToPoints(1.0),
        _ => Units.PixelsToPoints(1.0),
    };

    /// <summary>罫線スタイルに対応する破線パターン。</summary>
    public static LineDashStyle GetDash(BorderLineStyle style) => style switch
    {
        BorderLineStyle.Dotted => LineDashStyle.Dot,
        BorderLineStyle.Dashed => LineDashStyle.Dash,
        BorderLineStyle.MediumDashed => LineDashStyle.Dash,
        BorderLineStyle.DashDot => LineDashStyle.DashDot,
        BorderLineStyle.MediumDashDot => LineDashStyle.DashDot,
        BorderLineStyle.SlantDashDot => LineDashStyle.DashDot,
        BorderLineStyle.DashDotDot => LineDashStyle.DashDotDot,
        BorderLineStyle.MediumDashDotDot => LineDashStyle.DashDotDot,
        _ => LineDashStyle.Solid,
    };

    /// <summary>二重線かどうか。</summary>
    public static bool IsDouble(BorderLineStyle style) => style == BorderLineStyle.Double;
}
