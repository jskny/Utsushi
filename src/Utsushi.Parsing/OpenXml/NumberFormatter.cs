using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Utsushi.Parsing.OpenXml;

/// <summary>
/// Excel の数値書式(<c>numFmt</c>)を表示文字列へ適用する。
/// </summary>
/// <remarks>
/// <para>
/// 汎用のExcel数値書式エンジンではなく、<b>自社帳票が使用する範囲</b>(日付・金額・数量・パーセント)
/// を対象としたサブセット実装である(`.kiro/steering/product.md`「非対応(スコープ外)」)。
/// </para>
/// <para>
/// 解釈できない書式に遭遇した場合は例外にせず、値をそのまま(General 相当で)返す。
/// 表示が崩れた場合はゴールデンテストで検出し、必要に応じて対応パターンを追加する方針とする。
/// </para>
/// </remarks>
internal static class NumberFormatter
{
    /// <summary>Excel のシリアル値の基準日(1900年日付システム)。</summary>
    private static readonly DateTime SerialEpoch = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Unspecified);

    /// <summary>
    /// 組み込み数値書式ID(<c>numFmtId</c> 0〜49)のうち、対象帳票で現れうるものの書式文字列。
    /// ECMA-376 Part 1, 18.8.30 の既定値。
    /// </summary>
    private static readonly Dictionary<int, string> BuiltInFormats = new()
    {
        [0] = "General",
        [1] = "0",
        [2] = "0.00",
        [3] = "#,##0",
        [4] = "#,##0.00",
        [9] = "0%",
        [10] = "0.00%",
        [11] = "0.00E+00",
        [12] = "# ?/?",
        [13] = "# ??/??",
        [14] = "yyyy/mm/dd",
        [15] = "d-mmm-yy",
        [16] = "d-mmm",
        [17] = "mmm-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "yyyy/mm/dd h:mm",
        [37] = "#,##0;-#,##0",
        [38] = "#,##0;[Red]-#,##0",
        [39] = "#,##0.00;-#,##0.00",
        [40] = "#,##0.00;[Red]-#,##0.00",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mm:ss.0",
        [48] = "##0.0E+0",
        [49] = "@",
    };

    /// <summary>組み込み書式IDから書式文字列を取得する。未知のIDは null。</summary>
    public static string? GetBuiltInFormat(int numberFormatId) =>
        BuiltInFormats.TryGetValue(numberFormatId, out var format) ? format : null;

    /// <summary>
    /// 数値セルの値に書式を適用して表示文字列を返す。
    /// </summary>
    /// <param name="value">セルの生の数値。</param>
    /// <param name="formatCode">数値書式文字列。null / "General" の場合は既定表記。</param>
    public static string FormatNumber(double value, string? formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode) || IsGeneral(formatCode!))
        {
            return FormatGeneral(value);
        }

        var section = SelectSection(formatCode!, value);
        if (section is null || IsGeneral(section))
        {
            return FormatGeneral(value);
        }

        // 負数セクションが選ばれた場合、値の絶対値に対して書式を適用する
        // (符号はセクション側のリテラル "-" が担うため)。
        var target = SectionCountOf(formatCode!) > 1 && value < 0 ? Math.Abs(value) : value;

        try
        {
            return IsDateTimeFormat(section)
                ? FormatDateTime(target, section)
                : FormatNumericSection(target, section);
        }
        catch (Exception ex) when (ex is FormatException or ArgumentException or OverflowException)
        {
            // 解釈できない書式は General 相当にフォールバックする。
            return FormatGeneral(value);
        }
    }

    /// <summary>文字列セルに書式(<c>@</c> 等)を適用する。現状は素通し。</summary>
    public static string FormatText(string value, string? formatCode) => value;

    private static bool IsGeneral(string formatCode) =>
        formatCode.Trim().Equals("General", StringComparison.OrdinalIgnoreCase);

    /// <summary>Excelの General 表記(最大11桁相当の丸め、末尾ゼロ無し)を近似する。</summary>
    private static string FormatGeneral(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        var rounded = Math.Round(value, 10, MidpointRounding.AwayFromZero);
        return rounded.ToString("0.##########", CultureInfo.InvariantCulture);
    }

    private static int SectionCountOf(string formatCode) => SplitSections(formatCode).Count;

    /// <summary>
    /// <c>正;負;ゼロ;文字列</c> のセクションから、値に対応するものを選ぶ。
    /// </summary>
    private static string? SelectSection(string formatCode, double value)
    {
        var sections = SplitSections(formatCode);
        if (sections.Count == 0)
        {
            return null;
        }

        if (sections.Count == 1)
        {
            return sections[0];
        }

        if (value > 0)
        {
            return sections[0];
        }

        if (value < 0)
        {
            return sections[1];
        }

        return sections.Count >= 3 ? sections[2] : sections[0];
    }

    /// <summary>引用符・角括弧内のセミコロンを無視してセクション分割する。</summary>
    private static List<string> SplitSections(string formatCode)
    {
        var sections = new List<string>(4);
        var sb = new StringBuilder();
        var inQuotes = false;
        var inBrackets = false;

        for (var i = 0; i < formatCode.Length; i++)
        {
            var c = formatCode[i];
            switch (c)
            {
                case '"':
                    inQuotes = !inQuotes;
                    sb.Append(c);
                    break;
                case '[' when !inQuotes:
                    inBrackets = true;
                    sb.Append(c);
                    break;
                case ']' when !inQuotes:
                    inBrackets = false;
                    sb.Append(c);
                    break;
                case '\\' when i + 1 < formatCode.Length:
                    sb.Append(c).Append(formatCode[i + 1]);
                    i++;
                    break;
                case ';' when !inQuotes && !inBrackets:
                    sections.Add(sb.ToString());
                    sb.Clear();
                    break;
                default:
                    sb.Append(c);
                    break;
            }
        }

        sections.Add(sb.ToString());
        return sections;
    }

    /// <summary>日付/時刻の書式指定子を含むかどうかを判定する。</summary>
    internal static bool IsDateTimeFormat(string section)
    {
        var inQuotes = false;
        for (var i = 0; i < section.Length; i++)
        {
            var c = section[i];
            if (c == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (inQuotes)
            {
                continue;
            }

            if (c == '\\')
            {
                i++;
                continue;
            }

            if (c == '[')
            {
                // [Red] 等の色指定・条件指定は日付判定の対象外
                var close = section.IndexOf(']', i);
                i = close < 0 ? section.Length : close;
                continue;
            }

            if (c is 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 's' or 'S' or 'm' or 'M' or 'e' or 'g' or 'a' or 'A')
            {
                // 'm' は「月」と「分」の両方に使われるが、いずれにせよ日付時刻書式である。
                // ただし 'a'/'g' は AM/PM・元号以外では現れないため、前後関係は見ない。
                if (c is 'a' or 'A' && !LooksLikeAmPm(section, i))
                {
                    continue;
                }

                return true;
            }
        }

        return false;
    }

    private static bool LooksLikeAmPm(string section, int index) =>
        section.IndexOf("AM/PM", index, StringComparison.OrdinalIgnoreCase) == index
        || section.IndexOf("aaa", index, StringComparison.OrdinalIgnoreCase) == index;

    private static string FormatDateTime(double serial, string section)
    {
        var dateTime = FromSerial(serial);
        var sb = new StringBuilder();
        var culture = CultureInfo.GetCultureInfo("ja-JP");
        var hasAmPm = section.IndexOf("AM/PM", StringComparison.OrdinalIgnoreCase) >= 0;

        for (var i = 0; i < section.Length;)
        {
            var c = section[i];

            if (c == '"')
            {
                var end = section.IndexOf('"', i + 1);
                if (end < 0)
                {
                    sb.Append(section, i + 1, section.Length - i - 1);
                    break;
                }

                sb.Append(section, i + 1, end - i - 1);
                i = end + 1;
                continue;
            }

            if (c == '\\' && i + 1 < section.Length)
            {
                sb.Append(section[i + 1]);
                i += 2;
                continue;
            }

            if (c == '[')
            {
                var end = section.IndexOf(']', i);
                i = end < 0 ? section.Length : end + 1;
                continue;
            }

            if (c is '_' or '*')
            {
                // _x は x の幅の空白、*x は繰り返し。帳票の桁揃え用途では空白1つで近似する。
                i += i + 1 < section.Length ? 2 : 1;
                sb.Append(' ');
                continue;
            }

            var token = ReadRun(section, i, out var length);
            i += length;

            switch (token[0])
            {
                case 'y':
                case 'Y':
                    sb.Append(token.Length <= 2
                        ? (dateTime.Year % 100).ToString("00", CultureInfo.InvariantCulture)
                        : dateTime.Year.ToString("0000", CultureInfo.InvariantCulture));
                    break;

                case 'm':
                case 'M':
                    AppendMonthOrMinute(sb, section, i - length, token, dateTime, culture);
                    break;

                case 'd':
                case 'D':
                    sb.Append(token.Length switch
                    {
                        1 => dateTime.Day.ToString(CultureInfo.InvariantCulture),
                        2 => dateTime.Day.ToString("00", CultureInfo.InvariantCulture),
                        3 => culture.DateTimeFormat.AbbreviatedDayNames[(int)dateTime.DayOfWeek],
                        _ => culture.DateTimeFormat.DayNames[(int)dateTime.DayOfWeek],
                    });
                    break;

                case 'a':
                case 'A':
                    if (token.Equals("AM/PM", StringComparison.OrdinalIgnoreCase))
                    {
                        sb.Append(dateTime.Hour < 12 ? "AM" : "PM");
                    }
                    else if (token.StartsWith("aaaa", StringComparison.OrdinalIgnoreCase))
                    {
                        sb.Append(culture.DateTimeFormat.DayNames[(int)dateTime.DayOfWeek]);
                    }
                    else
                    {
                        sb.Append(culture.DateTimeFormat.AbbreviatedDayNames[(int)dateTime.DayOfWeek]);
                    }

                    break;

                case 'h':
                case 'H':
                {
                    var hour = hasAmPm ? ToTwelveHour(dateTime.Hour) : dateTime.Hour;
                    sb.Append(token.Length <= 1
                        ? hour.ToString(CultureInfo.InvariantCulture)
                        : hour.ToString("00", CultureInfo.InvariantCulture));
                    break;
                }

                case 's':
                case 'S':
                    sb.Append(token.Length <= 1
                        ? dateTime.Second.ToString(CultureInfo.InvariantCulture)
                        : dateTime.Second.ToString("00", CultureInfo.InvariantCulture));
                    break;

                default:
                    sb.Append(token);
                    break;
            }
        }

        return sb.ToString();
    }

    private static int ToTwelveHour(int hour)
    {
        var h = hour % 12;
        return h == 0 ? 12 : h;
    }

    /// <summary>
    /// 'm' は直前が h/[h] または直後が 's' の場合「分」、それ以外は「月」を表す。
    /// </summary>
    private static void AppendMonthOrMinute(
        StringBuilder sb, string section, int tokenStart, string token, DateTime dateTime, CultureInfo culture)
    {
        if (IsMinuteContext(section, tokenStart, token.Length))
        {
            sb.Append(token.Length <= 1
                ? dateTime.Minute.ToString(CultureInfo.InvariantCulture)
                : dateTime.Minute.ToString("00", CultureInfo.InvariantCulture));
            return;
        }

        sb.Append(token.Length switch
        {
            1 => dateTime.Month.ToString(CultureInfo.InvariantCulture),
            2 => dateTime.Month.ToString("00", CultureInfo.InvariantCulture),
            3 => culture.DateTimeFormat.AbbreviatedMonthNames[dateTime.Month - 1],
            4 => culture.DateTimeFormat.MonthNames[dateTime.Month - 1],
            _ => culture.DateTimeFormat.MonthNames[dateTime.Month - 1].Substring(0, 1),
        });
    }

    private static bool IsMinuteContext(string section, int tokenStart, int tokenLength)
    {
        for (var i = tokenStart - 1; i >= 0; i--)
        {
            var c = section[i];
            if (c is ' ' or ':' or ']')
            {
                continue;
            }

            if (c is 'h' or 'H')
            {
                return true;
            }

            break;
        }

        for (var i = tokenStart + tokenLength; i < section.Length; i++)
        {
            var c = section[i];
            if (c is ' ' or ':')
            {
                continue;
            }

            return c is 's' or 'S';
        }

        return false;
    }

    /// <summary>同じ文字の連続(AM/PM は特別扱い)を1トークンとして読む。</summary>
    private static string ReadRun(string section, int start, out int length)
    {
        if (section.IndexOf("AM/PM", start, StringComparison.OrdinalIgnoreCase) == start)
        {
            length = 5;
            return section.Substring(start, 5);
        }

        var c = section[start];
        var end = start + 1;
        while (end < section.Length && char.ToLowerInvariant(section[end]) == char.ToLowerInvariant(c))
        {
            end++;
        }

        length = end - start;
        return section.Substring(start, length);
    }

    /// <summary>Excel のシリアル値を <see cref="DateTime"/> へ変換する(1900年日付システム)。</summary>
    internal static DateTime FromSerial(double serial)
    {
        // Excel は 1900年をうるう年とみなす既知の不具合があり、シリアル値60が存在しない日付
        // (1900-02-29)に割り当てられている。60未満は1日ずらして補正する。
        var adjusted = serial < 60 ? serial + 1 : serial;
        return SerialEpoch.AddDays(adjusted);
    }

    private static string FormatNumericSection(double value, string section)
    {
        var parsed = NumericSection.Parse(section);
        return parsed.Format(value);
    }
}
