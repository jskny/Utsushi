namespace Utsushi.Core;

/// <summary>
/// 単位換算のユーティリティ。
/// </summary>
/// <remarks>
/// <para>
/// Utsushi の内部座標・寸法の単位は <b>ポイント(pt, 1pt = 1/72インチ)</b> に統一する。
/// 帳票定義や外部仕様でミリ単位を扱う場合も、レイヤーの境界を越える前にポイントへ変換する。
/// (`.kiro/steering/tech.md`「コーディング規約」、`design.md`「単位と座標系」)
/// </para>
/// <para>
/// Excel の一部の値は歴史的に96dpiのピクセルを前提としているため、
/// ピクセル⇔ポイントの換算にも 96dpi を用いる。
/// </para>
/// </remarks>
public static class Units
{
    /// <summary>1インチあたりのポイント数。</summary>
    public const double PointsPerInch = 72.0;

    /// <summary>1インチあたりのミリメートル数。</summary>
    public const double MillimetersPerInch = 25.4;

    /// <summary>Excel が前提とする画面解像度(dpi)。列幅のピクセル換算に用いる。</summary>
    public const double ExcelDpi = 96.0;

    /// <summary>1 EMU(English Metric Unit)あたりのポイント数。</summary>
    public const double EmusPerPoint = 12700.0;

    public static double InchesToPoints(double inches) => inches * PointsPerInch;

    public static double PointsToInches(double points) => points / PointsPerInch;

    public static double MillimetersToPoints(double millimeters) => millimeters / MillimetersPerInch * PointsPerInch;

    public static double PointsToMillimeters(double points) => points / PointsPerInch * MillimetersPerInch;

    /// <summary>96dpi のピクセル数をポイントへ換算する。</summary>
    public static double PixelsToPoints(double pixels) => pixels * PointsPerInch / ExcelDpi;

    /// <summary>ポイントを96dpiのピクセル数へ換算する。</summary>
    public static double PointsToPixels(double points) => points * ExcelDpi / PointsPerInch;

    public static double EmusToPoints(long emus) => emus / EmusPerPoint;
}
