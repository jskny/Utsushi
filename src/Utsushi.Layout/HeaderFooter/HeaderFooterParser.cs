using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Utsushi.Core;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout.HeaderFooter
{
    /// <summary>ヘッダー/フッターのセクション(横方向の配置)。</summary>
    public enum HeaderFooterSection
    {
        Left = 0,
        Center,
        Right,
    }

    /// <summary>同一の書式で連続する文字列。</summary>
    /// <param name="Text">展開済みの文字列(書式コードは実際の値に置換済み)。</param>
    /// <param name="Font">この部分のフォント。</param>
    public sealed record HeaderFooterRun(string Text, FontStyle Font);

    /// <summary>1セクション分の内容。</summary>
    /// <param name="Section">配置。</param>
    /// <param name="Runs">書式ごとに分割された文字列。</param>
    public sealed record HeaderFooterPart(HeaderFooterSection Section, IReadOnlyList<HeaderFooterRun> Runs)
    {
        public bool IsEmpty
        {
            get
            {
                foreach (var run in Runs)
                {
                    if (run.Text.Length > 0)
                    {
                        return false;
                    }
                }

                return true;
            }
        }
    }

    /// <summary>
    /// ヘッダー/フッターの書式コード展開に必要な文脈。
    /// </summary>
    /// <param name="PageNumber">1始まりのページ番号(<c>&amp;P</c>)。</param>
    /// <param name="TotalPages">総ページ数(<c>&amp;N</c>)。</param>
    /// <param name="SheetName">シート名(<c>&amp;A</c>)。</param>
    /// <param name="FileName">ファイル名(<c>&amp;F</c>)。</param>
    /// <param name="Timestamp">日付・時刻(<c>&amp;D</c> / <c>&amp;T</c>)。</param>
    /// <param name="DefaultFont">書式コードで指定が無い場合に使うフォント。</param>
    public sealed record HeaderFooterContext(
        int PageNumber,
        int TotalPages,
        string SheetName,
        string? FileName,
        DateTime Timestamp,
        FontStyle DefaultFont);

    /// <summary>
    /// Excel のヘッダー/フッター書式コードを解釈し、セクションごとの文字列へ展開する。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 対応する書式コードは自社帳票で使う範囲に限定する(要件3.7, 3.8 の補足)。
    /// </para>
    /// <list type="table">
    ///   <item><term>&amp;L / &amp;C / &amp;R</term><description>以降を左/中央/右セクションへ</description></item>
    ///   <item><term>&amp;P / &amp;N</term><description>ページ番号 / 総ページ数</description></item>
    ///   <item><term>&amp;D / &amp;T</term><description>日付 / 時刻</description></item>
    ///   <item><term>&amp;A / &amp;F / &amp;Z</term><description>シート名 / ファイル名 / ファイルパス</description></item>
    ///   <item><term>&amp;B / &amp;I / &amp;U / &amp;S</term><description>太字 / 斜体 / 下線 / 取り消し線の切り替え</description></item>
    ///   <item><term>&amp;"フォント名,スタイル"</term><description>フォントの切り替え</description></item>
    ///   <item><term>&amp;nn(数字)</term><description>フォントサイズの切り替え</description></item>
    ///   <item><term>&amp;Krrggbb</term><description>文字色の指定</description></item>
    ///   <item><term>&amp;&amp;</term><description>文字としての &amp;</description></item>
    /// </list>
    /// <para>
    /// 上記以外の書式コード(画像の <c>&amp;G</c> など)は読み飛ばす。
    /// 未知のコードで例外にはせず、その帳票で必要になった時点で対応を追加する方針とする。
    /// </para>
    /// </remarks>
    public static class HeaderFooterParser
    {
        /// <summary>書式コードを解釈し、内容を持つセクションのみを返す。</summary>
        public static IReadOnlyList<HeaderFooterPart> Parse(string? definition, HeaderFooterContext context)
        {
            var parts = new List<HeaderFooterPart>(3);
            if (string.IsNullOrEmpty(definition))
            {
                return parts;
            }

            var builder = new SectionBuilder(context.DefaultFont);

            // セクション指定が無い場合、Excel は中央セクション扱いにする。
            builder.SwitchSection(HeaderFooterSection.Center);

            var text = definition!;
            for (var i = 0; i < text.Length;)
            {
                var c = text[i];
                if (c != '&')
                {
                    builder.Append(c);
                    i++;
                    continue;
                }

                if (i + 1 >= text.Length)
                {
                    // 末尾の単独の & は文字として扱う。
                    builder.Append(c);
                    break;
                }

                i = HandleCode(text, i, builder, context);
            }

            foreach (var part in builder.Build())
            {
                if (!part.IsEmpty)
                {
                    parts.Add(part);
                }
            }

            return parts;
        }

        /// <summary>書式コードを1つ処理し、次に読むべき位置を返す。</summary>
        private static int HandleCode(string text, int index, SectionBuilder builder, HeaderFooterContext context)
        {
            var code = text[index + 1];

            switch (code)
            {
                case '&':
                    builder.Append('&');
                    return index + 2;

                case 'L' or 'l':
                    builder.SwitchSection(HeaderFooterSection.Left);
                    return index + 2;

                case 'C' or 'c':
                    builder.SwitchSection(HeaderFooterSection.Center);
                    return index + 2;

                case 'R' or 'r':
                    builder.SwitchSection(HeaderFooterSection.Right);
                    return index + 2;

                case 'P' or 'p':
                    builder.Append(context.PageNumber.ToString(CultureInfo.InvariantCulture));
                    return index + 2;

                case 'N' or 'n':
                    builder.Append(context.TotalPages.ToString(CultureInfo.InvariantCulture));
                    return index + 2;

                case 'D' or 'd':
                    builder.Append(context.Timestamp.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture));
                    return index + 2;

                case 'T' or 't':
                    builder.Append(context.Timestamp.ToString("H:mm", CultureInfo.InvariantCulture));
                    return index + 2;

                case 'A' or 'a':
                    builder.Append(context.SheetName);
                    return index + 2;

                case 'F' or 'f':
                case 'Z' or 'z':
                    builder.Append(context.FileName ?? string.Empty);
                    return index + 2;

                case 'B' or 'b':
                    builder.ToggleBold();
                    return index + 2;

                case 'I' or 'i':
                    builder.ToggleItalic();
                    return index + 2;

                case 'U' or 'u':
                    builder.ToggleUnderline();
                    return index + 2;

                case 'S' or 's':
                    builder.ToggleStrike();
                    return index + 2;

                case '"':
                    return ApplyFontName(text, index, builder);

                case 'K' or 'k':
                    return ApplyFontColor(text, index, builder);

                case 'G' or 'g':
                    // 画像は対象外。
                    return index + 2;

                default:
                    if (char.IsDigit(code))
                    {
                        return ApplyFontSize(text, index, builder);
                    }

                    // 未知のコードは読み飛ばす。
                    return index + 2;
            }
        }

        /// <summary>&amp;"フォント名,スタイル" を処理する。</summary>
        private static int ApplyFontName(string text, int index, SectionBuilder builder)
        {
            var close = text.IndexOf('"', index + 2);
            if (close < 0)
            {
                return text.Length;
            }

            var spec = text.Substring(index + 2, close - index - 2);
            var comma = spec.IndexOf(',');
            var name = comma < 0 ? spec : spec.Substring(0, comma);
            var style = comma < 0 ? string.Empty : spec.Substring(comma + 1);

            // フォント名 "-" は「既定のフォント」を意味する。
            if (!string.IsNullOrEmpty(name) && name != "-")
            {
                builder.SetFontName(name);
            }

            builder.SetFontStyleFromName(style);
            return close + 1;
        }

        /// <summary>
        /// &amp;Krrggbb(文字色)を処理する。<paramref name="text"/>の<paramref name="index"/>+2から
        /// 6桁を色として解釈できた場合のみフォント色に反映する。桁数不足・16進以外の文字が
        /// 含まれる場合は既定色のまま変更しない(例外にはしない)。いずれの場合も6桁ぶんは
        /// 読み飛ばす(Excel自身も不正な値をそのまま消費する挙動に合わせる)。
        /// </summary>
        private static int ApplyFontColor(string text, int index, SectionBuilder builder)
        {
            var end = Math.Min(text.Length, index + 2 + 6);
            var hex = text.Substring(index + 2, end - (index + 2));
            if (hex.Length == 6 && ArgbColor.TryParseHex(hex, out var color))
            {
                builder.SetFontColor(color);
            }

            return end;
        }

        /// <summary>&amp;nn(数字)を処理する。</summary>
        private static int ApplyFontSize(string text, int index, SectionBuilder builder)
        {
            var start = index + 1;
            var end = start;
            while (end < text.Length && char.IsDigit(text[end]))
            {
                end++;
            }

            var digits = text.Substring(start, end - start);
            if (double.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out var size) && size > 0)
            {
                builder.SetFontSize(size);
            }

            return end;
        }

        /// <summary>セクションごとに、書式の変わり目で文字列を区切りながら組み立てる。</summary>
        private sealed class SectionBuilder
        {
            private readonly FontStyle _defaultFont;
            private readonly Dictionary<HeaderFooterSection, List<HeaderFooterRun>> _runs = new();
            private readonly StringBuilder _current = new();

            private HeaderFooterSection _section = HeaderFooterSection.Center;
            private FontStyle _font;

            public SectionBuilder(FontStyle defaultFont)
            {
                _defaultFont = defaultFont;
                _font = defaultFont;

                foreach (HeaderFooterSection section in Enum.GetValues(typeof(HeaderFooterSection)))
                {
                    _runs[section] = new List<HeaderFooterRun>();
                }
            }

            public void Append(char c) => _current.Append(c);

            public void Append(string text) => _current.Append(text);

            public void SwitchSection(HeaderFooterSection section)
            {
                Flush();
                _section = section;

                // Excel はセクションが変わるとフォント指定を引き継がず既定に戻す。
                _font = _defaultFont;
            }

            public void ToggleBold() => ChangeFont(_font with { Bold = !_font.Bold });

            public void ToggleItalic() => ChangeFont(_font with { Italic = !_font.Italic });

            public void ToggleUnderline() => ChangeFont(_font with
            {
                Underline = _font.Underline == UnderlineStyle.None ? UnderlineStyle.Single : UnderlineStyle.None,
            });

            public void ToggleStrike() => ChangeFont(_font with { Strike = !_font.Strike });

            public void SetFontName(string name) => ChangeFont(_font with { Name = name });

            public void SetFontColor(ArgbColor color) => ChangeFont(_font with { Color = color });

            public void SetFontSize(double sizePt) => ChangeFont(_font with { SizePt = sizePt });

            /// <summary>&amp;"フォント名,スタイル" のスタイル部分(Bold / Italic / Bold Italic / Regular)を反映する。</summary>
            public void SetFontStyleFromName(string style)
            {
                if (string.IsNullOrWhiteSpace(style))
                {
                    return;
                }

                var normalized = style.Trim().ToLowerInvariant();
                var bold = normalized.Contains("bold");
                var italic = normalized.Contains("italic") || normalized.Contains("oblique");
                ChangeFont(_font with { Bold = bold, Italic = italic });
            }

            public IReadOnlyList<HeaderFooterPart> Build()
            {
                Flush();

                var parts = new List<HeaderFooterPart>(_runs.Count);
                foreach (var section in new[]
                         {
                         HeaderFooterSection.Left, HeaderFooterSection.Center, HeaderFooterSection.Right,
                     })
                {
                    parts.Add(new HeaderFooterPart(section, _runs[section]));
                }

                return parts;
            }

            private void ChangeFont(FontStyle font)
            {
                Flush();
                _font = font;
            }

            private void Flush()
            {
                if (_current.Length == 0)
                {
                    return;
                }

                _runs[_section].Add(new HeaderFooterRun(_current.ToString(), _font));
                _current.Clear();
            }
        }
    }
}
