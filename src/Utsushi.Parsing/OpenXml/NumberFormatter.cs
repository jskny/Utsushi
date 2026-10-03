using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Utsushi.Core;

namespace Utsushi.Parsing.OpenXml
{
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

        /// <summary>1日あたりのミリ秒数。</summary>
        private const double MillisecondsPerDay = 24.0 * 60.0 * 60.0 * 1000.0;

        /// <summary>和暦の短い日付(例: <c>R8.4.1</c>)。組み込み書式 27・36・50・57。</summary>
        private const string JapaneseEraShortDate = "[$-411]ge.m.d";

        /// <summary>和暦の長い日付(例: <c>令和8年4月1日</c>)。組み込み書式 28・29・51・54・58。</summary>
        private const string JapaneseEraLongDate = "[$-411]ggge\"年\"m\"月\"d\"日\"";

        /// <summary>Excel の書式コードの最大長。これを超える書式コードは解析結果をキャッシュしない。</summary>
        private const int MaxCacheableFormatLength = 255;

        /// <summary>
        /// 解析結果キャッシュの件数上限。上限に達したらキャッシュを空にして登録し直す(先着の書式で埋まったままだと、
        /// 常駐プロセスでは1回の入力で埋められた後、以降の変換の書式がすべて都度解析になる。security-reviewer指摘)。
        /// </summary>
        private const int MaxCacheEntries = 1024;

        /// <summary>書式コードごとの解析結果。数値セルごとの再解析を避ける。</summary>
        private static readonly ConcurrentDictionary<string, ParsedFormat> ParsedFormats = new(StringComparer.Ordinal);

        /// <summary>
        /// 元号の一覧(開始日の昇順)。明治の開始日は Excel の扱いに合わせて 1868/1/1 とする。
        /// </summary>
        private static readonly JapaneseEra[] JapaneseEras =
        {
            new(new DateTime(1868, 1, 1), 'M', "明", "明治"),
            new(new DateTime(1912, 7, 30), 'T', "大", "大正"),
            new(new DateTime(1926, 12, 25), 'S', "昭", "昭和"),
            new(new DateTime(1989, 1, 8), 'H', "平", "平成"),
            new(new DateTime(2019, 5, 1), 'R', "令", "令和"),
        };

        /// <summary>
        /// 組み込み数値書式ID(<c>numFmtId</c> 0〜58)のうち、対象帳票で現れうるものの書式文字列。
        /// </summary>
        /// <remarks>
        /// 本製品は日本の帳票が対象のため、ECMA-376 Part 1, 18.8.30 の既定値のうち
        /// 日本語(ja-JP)ロケールの表示に合わせている(例: 14 は <c>yyyy/m/d</c>、5〜8 は円記号、
        /// 27〜36・50〜58 は和暦を含む日本語ロケール固有の日付・時刻書式)。
        /// </remarks>
        private static readonly Dictionary<int, string> BuiltInFormats = new()
        {
            [0] = "General",
            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [4] = "#,##0.00",
            [5] = "\"¥\"#,##0;\"¥\"\\-#,##0",
            [6] = "\"¥\"#,##0;[Red]\"¥\"\\-#,##0",
            [7] = "\"¥\"#,##0.00;\"¥\"\\-#,##0.00",
            [8] = "\"¥\"#,##0.00;[Red]\"¥\"\\-#,##0.00",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00",
            [12] = "# ?/?",
            [13] = "# ??/??",
            [14] = "yyyy/m/d",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "yyyy/m/d h:mm",
            [27] = JapaneseEraShortDate,
            [28] = JapaneseEraLongDate,
            [29] = JapaneseEraLongDate,
            [30] = "m/d/yy",
            [31] = "yyyy\"年\"m\"月\"d\"日\"",
            [32] = "h\"時\"mm\"分\"",
            [33] = "h\"時\"mm\"分\"ss\"秒\"",
            [34] = "yyyy\"年\"m\"月\"",
            [35] = "m\"月\"d\"日\"",
            [36] = JapaneseEraShortDate,
            [37] = "#,##0;-#,##0",
            [38] = "#,##0;[Red]-#,##0",
            [39] = "#,##0.00;-#,##0.00",
            [40] = "#,##0.00;[Red]-#,##0.00",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mm:ss.0",
            [48] = "##0.0E+0",
            [49] = "@",
            [50] = JapaneseEraShortDate,
            [51] = JapaneseEraLongDate,
            [52] = "yyyy\"年\"m\"月\"",
            [53] = "m\"月\"d\"日\"",
            [54] = JapaneseEraLongDate,
            [55] = "yyyy\"年\"m\"月\"",
            [56] = "m\"月\"d\"日\"",
            [57] = JapaneseEraShortDate,
            [58] = JapaneseEraLongDate,
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

            var parsed = GetParsedFormat(formatCode!);
            var section = parsed.Select(value);
            if (section is null || section.Kind == SectionKind.General)
            {
                // "@"(文字列書式)だけの書式など、数値に使えるセクションが無い場合も General で表示する。
                return FormatGeneral(value);
            }

            // 解釈できない書式は General 相当にフォールバックする。
            return TryFormatSection(value, parsed, section, out var text) ? text : FormatGeneral(value);
        }

        /// <summary>選んだセクションで書式を適用する。解釈できない書式・書式で表せない値なら false。</summary>
        private static bool TryFormatSection(double value, ParsedFormat parsed, FormatSection section, out string text)
        {
            // セクションが複数あり負数セクションが選ばれた場合、値の絶対値に対して書式を適用する
            // (符号はセクション側のリテラル "-" が担うため)。
            // セクションが1つだけの場合は負号付きのまま渡し、NumericSection が出力全体の先頭に負号を付ける。
            var target = parsed.NumericSections.Count > 1 && value < 0 ? Math.Abs(value) : value;

            try
            {
                text = section.Kind == SectionKind.DateTime
                    ? FormatDateTime(target, section.Text)
                    : section.Numeric!.Format(target);
                return true;
            }
            catch (Exception ex) when (ex is FormatException or ArgumentException or OverflowException)
            {
                text = string.Empty;
                return false;
            }
        }

        /// <summary>
        /// 数値に書式を適用するときに使うセクションの色の指定(要件4.12)を返す。色の指定が無い、
        /// または <see cref="FormatNumber"/> が General で表示する場合は null。
        /// </summary>
        /// <param name="value">セルの生の数値。</param>
        /// <param name="formatCode">数値書式文字列。</param>
        public static ArgbColor? ResolveColor(double value, string? formatCode)
        {
            if (string.IsNullOrWhiteSpace(formatCode) || IsGeneral(formatCode!))
            {
                return null;
            }

            var parsed = GetParsedFormat(formatCode!);
            var section = parsed.Select(value);
            if (section?.Color is not { } color || section.Kind == SectionKind.General)
            {
                return null;
            }

            // FormatNumber が General の表示に戻す値(負の経過時間・明治より前の和暦など)は、色も付けない。
            return TryFormatSection(value, parsed, section, out _) ? color : null;
        }

        /// <summary>
        /// セクション内の角括弧の指定から色を探す。英語の色名・日本語版Excelの色名・<c>[ColorN]</c>(1〜56)を解釈し、
        /// それ以外(ロケール <c>[$-411]</c>・経過時間 <c>[h]</c>・条件 <c>[&gt;100]</c> など)は無視する。
        /// </summary>
        private static ArgbColor? ParseSectionColor(string section)
        {
            var inQuotes = false;
            for (var i = 0; i < section.Length; i++)
            {
                var c = section[i];
                if (c == '\\')
                {
                    i++;
                    continue;
                }

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (inQuotes || c != '[')
                {
                    continue;
                }

                var end = section.IndexOf(']', i + 1);
                if (end < 0)
                {
                    return null;
                }

                if (TryParseColorName(section.Substring(i + 1, end - i - 1), out var color))
                {
                    return color;
                }

                i = end;
            }

            return null;
        }

        private static bool TryParseColorName(string name, out ArgbColor color)
        {
            uint? argb = name.ToUpperInvariant() switch
            {
                "BLACK" or "黒" => 0xFF000000,
                "WHITE" or "白" => 0xFFFFFFFF,
                "RED" or "赤" => 0xFFFF0000,
                "GREEN" or "緑" => 0xFF00FF00,
                "BLUE" or "青" => 0xFF0000FF,
                "YELLOW" or "黄" => 0xFFFFFF00,
                "MAGENTA" or "紫" => 0xFFFF00FF,
                "CYAN" or "水" => 0xFF00FFFF,
                _ => null,
            };

            if (argb is null
                && name.StartsWith("COLOR", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(name.Substring(5), NumberStyles.None, CultureInfo.InvariantCulture, out var index)
                && index >= 1 && index <= 56)
            {
                // [ColorN] はインデックスカラーの N+7 番(Excel の既定のパレット)。
                argb = ColorResolver.DefaultIndexedPalette[index + 7];
            }

            color = argb is { } v
                ? new ArgbColor((byte)(v >> 24), (byte)(v >> 16), (byte)(v >> 8), (byte)v)
                : default;
            return argb is not null;
        }

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

        /// <summary>書式コードの解析結果をキャッシュから取得する。無ければ解析して(上限内なら)登録する。</summary>
        private static ParsedFormat GetParsedFormat(string formatCode)
        {
            if (ParsedFormats.TryGetValue(formatCode, out var cached))
            {
                return cached;
            }

            var parsed = ParsedFormat.Parse(formatCode);
            if (formatCode.Length <= MaxCacheableFormatLength)
            {
                if (ParsedFormats.Count >= MaxCacheEntries)
                {
                    ParsedFormats.Clear();
                }

                // 並行して同じ書式を解析した場合はどちらか一方が登録される(内容は同一)。
                ParsedFormats.TryAdd(formatCode, parsed);
            }

            return parsed;
        }

        /// <summary>引用符・角括弧・エスケープの外にある <c>@</c>(文字列の差し込み位置)を含むかどうか。</summary>
        private static bool ContainsTextPlaceholder(string section)
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
                    var close = section.IndexOf(']', i);
                    i = close < 0 ? section.Length : close;
                    continue;
                }

                if (c == '@')
                {
                    return true;
                }
            }

            return false;
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
                    // [h]・[mm]・[ss] は経過時間の書式指定子。
                    // [Red] 等の色指定・[$-411] 等のロケール指定・条件指定は日付判定の対象外。
                    var close = section.IndexOf(']', i);
                    if (close > i && IsElapsedTimeToken(section, i + 1, close - i - 1))
                    {
                        return true;
                    }

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

        /// <summary>
        /// 角括弧の中身が経過時間の書式指定子(<c>h</c>・<c>m</c>・<c>s</c> のいずれか1種類の連続)かどうか。
        /// </summary>
        private static bool IsElapsedTimeToken(string section, int start, int length)
        {
            if (length <= 0)
            {
                return false;
            }

            var first = char.ToLowerInvariant(section[start]);
            if (first is not ('h' or 'm' or 's'))
            {
                return false;
            }

            for (var k = 1; k < length; k++)
            {
                if (char.ToLowerInvariant(section[start + k]) != first)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool LooksLikeAmPm(string section, int index) =>
            section.IndexOf("AM/PM", index, StringComparison.OrdinalIgnoreCase) == index
            || section.IndexOf("aaa", index, StringComparison.OrdinalIgnoreCase) == index;

        private static bool IsAsciiLetter(char c) => c is (>= 'a' and <= 'z') or (>= 'A' and <= 'Z');

        private static string FormatDateTime(double serial, string section)
        {
            var dateTime = FromSerial(serial);
            var sb = new StringBuilder();
            var culture = CultureInfo.GetCultureInfo("ja-JP");
            var hasAmPm = section.IndexOf("AM/PM", StringComparison.OrdinalIgnoreCase) >= 0;
            var secondsEmitted = false;

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
                    if (end > i && IsElapsedTimeToken(section, i + 1, end - i - 1))
                    {
                        AppendElapsedTime(sb, serial, section[i + 1], end - i - 1);
                        secondsEmitted |= section[i + 1] is 's' or 'S';
                        i = end + 1;
                        continue;
                    }

                    // [Red] 等の色指定・[$-411] 等のロケール指定は表示に反映しない。
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

                if (c == '.' && secondsEmitted && i + 1 < section.Length && section[i + 1] == '0')
                {
                    // ss.0 / ss.00 / ss.000 は秒の小数部(最大3桁、切り捨て)。
                    var digits = 0;
                    while (digits < 3 && i + 1 + digits < section.Length && section[i + 1 + digits] == '0')
                    {
                        digits++;
                    }

                    var divisor = digits switch { 1 => 100, 2 => 10, _ => 1 };
                    sb.Append('.').Append(
                        (dateTime.Millisecond / divisor).ToString(new string('0', digits), CultureInfo.InvariantCulture));
                    i += 1 + digits;
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
                        secondsEmitted = true;
                        break;

                    case 'g':
                    case 'G':
                        {
                            // g: 英字1文字(R)、gg: 漢字1文字(令)、ggg: 元号名(令和)
                            var era = FindJapaneseEra(dateTime);
                            sb.Append(token.Length switch
                            {
                                1 => era.Letter.ToString(),
                                2 => era.Abbreviation,
                                _ => era.Name,
                            });
                            break;
                        }

                    case 'e':
                    case 'E':
                        {
                            // e: 元号の年(1年は「1」)、ee: 2桁
                            var eraYear = dateTime.Year - FindJapaneseEra(dateTime).Start.Year + 1;
                            sb.Append(token.Length <= 1
                                ? eraYear.ToString(CultureInfo.InvariantCulture)
                                : eraYear.ToString("00", CultureInfo.InvariantCulture));
                            break;
                        }

                    default:
                        if (IsAsciiLetter(token[0]))
                        {
                            // 解釈できない書式指定子(b 等)は、書式文字をそのまま出さず General 相当にフォールバックする。
                            throw new FormatException($"未対応の日付書式指定子 '{token}'。");
                        }

                        sb.Append(token);
                        break;
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// 経過時間 <c>[h]</c>・<c>[m]</c>・<c>[s]</c> を追記する(24時間・60分・60秒を超えて数える)。
        /// </summary>
        private static void AppendElapsedTime(StringBuilder sb, double serial, char unit, int width)
        {
            if (serial < 0)
            {
                throw new FormatException("負の経過時間は表示できない。");
            }

            // 時・分・秒の表示(FromSerial)と同じくミリ秒に丸めてから数える。
            var totalMilliseconds = (long)Math.Round(serial * MillisecondsPerDay, MidpointRounding.AwayFromZero);
            var elapsed = char.ToLowerInvariant(unit) switch
            {
                'h' => totalMilliseconds / 3_600_000L,
                'm' => totalMilliseconds / 60_000L,
                _ => totalMilliseconds / 1_000L,
            };

            sb.Append(elapsed.ToString(new string('0', width), CultureInfo.InvariantCulture));
        }

        /// <summary>日付が属する元号を返す。明治より前の日付は和暦で表せないため例外にする。</summary>
        private static JapaneseEra FindJapaneseEra(DateTime dateTime)
        {
            var date = dateTime.Date;
            for (var k = JapaneseEras.Length - 1; k >= 0; k--)
            {
                if (date >= JapaneseEras[k].Start)
                {
                    return JapaneseEras[k];
                }
            }

            throw new FormatException("明治より前の日付は和暦で表示できない。");
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

        /// <summary>
        /// 前後にある書式指定子(英字)を見て 'm' が「分」かどうかを判定する。
        /// 区切り記号・角括弧・引用符で囲んだリテラル(例: <c>h"時"mm"分"</c>)は読み飛ばす。
        /// </summary>
        private static bool IsMinuteContext(string section, int tokenStart, int tokenLength)
        {
            for (var i = tokenStart - 1; i >= 0; i--)
            {
                var c = section[i];
                if (c == '"')
                {
                    var open = i > 0 ? section.LastIndexOf('"', i - 1) : -1;
                    if (open < 0)
                    {
                        break;
                    }

                    i = open;
                    continue;
                }

                if (!IsAsciiLetter(c))
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
                if (c == '"')
                {
                    var close = section.IndexOf('"', i + 1);
                    if (close < 0)
                    {
                        break;
                    }

                    i = close;
                    continue;
                }

                if (!IsAsciiLetter(c))
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

            // シリアル値の小数部は時刻を表すが、double の丸め誤差でそのまま日数加算すると
            // 1ティック足りずに「9:05:00」が「9:04:59.999…」になることがある。
            // ミリ秒に丸めてから加算する(日数のままでは値が大きく、doubleの整数精度を超える)。
            var totalMilliseconds = Math.Round(adjusted * MillisecondsPerDay, MidpointRounding.AwayFromZero);
            return SerialEpoch.AddTicks((long)totalMilliseconds * TimeSpan.TicksPerMillisecond);
        }

        private enum SectionKind
        {
            General,
            DateTime,
            Numeric,
        }

        /// <summary>書式コード1セクション分の解析結果。</summary>
        private sealed class FormatSection
        {
            public FormatSection(string text)
            {
                Text = text;
                Color = ParseSectionColor(text);
                if (IsGeneral(text))
                {
                    Kind = SectionKind.General;
                }
                else if (IsDateTimeFormat(text))
                {
                    Kind = SectionKind.DateTime;
                }
                else
                {
                    Kind = SectionKind.Numeric;
                    Numeric = NumericSection.Parse(text);
                }
            }

            public string Text { get; }

            /// <summary>色の指定(要件4.12)。無ければ null。</summary>
            public ArgbColor? Color { get; }

            public SectionKind Kind { get; }

            /// <summary><see cref="Kind"/> が <see cref="SectionKind.Numeric"/> のときの解析結果。</summary>
            public NumericSection? Numeric { get; }
        }

        /// <summary>
        /// 書式コード全体の解析結果。不変であり、複数スレッドから共有してよい。
        /// </summary>
        private sealed class ParsedFormat
        {
            private ParsedFormat(IReadOnlyList<FormatSection> numericSections)
            {
                NumericSections = numericSections;
            }

            /// <summary>数値の表示に使うセクション(<c>@</c> を含む文字列用セクションを除く)。</summary>
            public IReadOnlyList<FormatSection> NumericSections { get; }

            public static ParsedFormat Parse(string formatCode)
            {
                var sections = new List<FormatSection>(4);
                foreach (var text in SplitSections(formatCode))
                {
                    // "@" を含むセクションは文字列の表示用で、数値の表示には使わない。
                    if (ContainsTextPlaceholder(text))
                    {
                        continue;
                    }

                    sections.Add(new FormatSection(text));
                }

                return new ParsedFormat(sections);
            }

            /// <summary><c>正;負;ゼロ</c> のセクションから、値に対応するものを選ぶ。</summary>
            public FormatSection? Select(double value)
            {
                var sections = NumericSections;
                if (sections.Count == 0)
                {
                    return null;
                }

                if (sections.Count == 1 || value > 0)
                {
                    return sections[0];
                }

                if (value < 0)
                {
                    return sections[1];
                }

                return sections.Count >= 3 ? sections[2] : sections[0];
            }
        }

        /// <summary>元号(開始日・英字略号・漢字略号・元号名)。</summary>
        private sealed class JapaneseEra
        {
            public JapaneseEra(DateTime start, char letter, string abbreviation, string name)
            {
                Start = start;
                Letter = letter;
                Abbreviation = abbreviation;
                Name = name;
            }

            public DateTime Start { get; }

            public char Letter { get; }

            public string Abbreviation { get; }

            public string Name { get; }
        }
    }
}
