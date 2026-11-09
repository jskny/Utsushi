using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// リッチテキスト(ランごとの書式)が、セルの書式と異なる見た目かどうかを判定する(excel2pdf-consolidation 要件7.2(d))。
    /// </summary>
    /// <remarks>
    /// Utsushi はリッチテキストをセル全体のセルの書式で描く。ランの書式(<c>rPr</c>)のうち、フォント名・サイズ・太字・斜体・下線・
    /// 取り消し線・色のいずれかがセルの書式と異なるか、それ以外の見た目に関わる指定(上付き・下付き・輪郭・影など)を持つランがあれば、
    /// Excel と見た目が異なるとみなす。文字セット・フォントの系統・スキーム(<c>charset</c>・<c>family</c>・<c>scheme</c>)は見た目に関わらないため比べない。
    /// </remarks>
    internal sealed class RichTextFormats
    {
        /// <summary>
        /// 1つの文字列について保持する、異なるランの書式の数の上限。超えたら、ランの書式がセルの書式と異なるとみなす
        /// (同じ共有文字列を参照するセルの数×ランの数の比較で、CPU・メモリを使わせないため)。
        /// </summary>
        internal const int MaxDistinctRunFormats = 16;

        /// <summary>異なるランの書式が上限を超えた文字列の印(どのセルの書式とも異なるとみなす)。</summary>
        private static readonly IReadOnlyList<RunFormat> TooManyFormats = Array.Empty<RunFormat>();

        private readonly IReadOnlyList<IReadOnlyList<RunFormat>?> _sharedStringRuns;

        private RichTextFormats(IReadOnlyList<IReadOnlyList<RunFormat>?> sharedStringRuns)
        {
            _sharedStringRuns = sharedStringRuns;
        }

        /// <summary>共有文字列ごとのランの書式を読む(<paramref name="count"/> 件まで。共有文字列の読み取りと同じ件数)。</summary>
        public static RichTextFormats Create(WorkbookPart workbookPart, ColorResolver colors, int count)
        {
            var result = new List<IReadOnlyList<RunFormat>?>();
            var table = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (table is not null)
            {
                foreach (var item in table.Elements<X.SharedStringItem>())
                {
                    if (result.Count >= count)
                    {
                        break;
                    }

                    result.Add(ReadRuns(item.Elements<X.Run>(), colors));
                }
            }

            return new RichTextFormats(result);
        }

        /// <summary>共有文字列 <paramref name="index"/> のランの書式が、セルのフォントと異なるかどうか。</summary>
        public bool SharedStringDiffers(int index, FontStyle cellFont) =>
            index >= 0 && index < _sharedStringRuns.Count && _sharedStringRuns[index] is { } runs && Differs(runs, cellFont);

        /// <summary>インライン文字列のランの書式が、セルのフォントと異なるかどうか。</summary>
        public static bool InlineDiffers(IEnumerable<X.Run> runs, FontStyle cellFont, ColorResolver colors) =>
            ReadRuns(runs, colors) is { } formats && Differs(formats, cellFont);

        /// <summary>
        /// ランの書式を、重複を除いて読む。書式を持つランが無ければ null。異なる書式が <see cref="MaxDistinctRunFormats"/> を
        /// 超えたら、読むのをやめて <see cref="TooManyFormats"/> を返す。
        /// </summary>
        private static IReadOnlyList<RunFormat>? ReadRuns(IEnumerable<X.Run> runs, ColorResolver colors)
        {
            HashSet<RunFormat>? seen = null;
            List<RunFormat>? formats = null;
            foreach (var run in runs)
            {
                if (run.RunProperties is not { } properties)
                {
                    continue;
                }

                var format = RunFormat.Read(properties, colors);
                if ((seen ??= new HashSet<RunFormat>()).Add(format))
                {
                    if (seen.Count > MaxDistinctRunFormats)
                    {
                        return TooManyFormats;
                    }

                    (formats ??= new List<RunFormat>()).Add(format);
                }
            }

            return formats;
        }

        private static bool Differs(IReadOnlyList<RunFormat> runs, FontStyle font)
        {
            if (ReferenceEquals(runs, TooManyFormats))
            {
                return true;
            }

            foreach (var run in runs)
            {
                if (run.Differs(font))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 1つのランの書式。フォント名・サイズ・色は指定のあるものだけ。太字・斜体・下線・取り消し線は、指定が無ければ無し
        /// (Excel はランの書式をすべて書き出すため、<c>rPr</c> に <c>b</c> が無いランは太字ではない)。
        /// </summary>
        /// <remarks>
        /// 診断にしか使わない値のため、不正な値はファイルの破損として扱わず、無視する(既定値とみなす)。
        /// </remarks>
        private sealed record RunFormat(
            string? Name, double? Size, bool Bold, bool Italic, bool Underline, bool Strike, ArgbColor? Color, bool HasOtherFormatting)
        {
            public static RunFormat Read(X.RunProperties properties, ColorResolver colors)
            {
                string? name = null;
                double? size = null;
                bool bold = false, italic = false, underline = false, strike = false;
                ArgbColor? color = null;
                var other = false;

                foreach (var element in properties.ChildElements)
                {
                    switch (element)
                    {
                        case X.RunFont runFont:
                            name = runFont.Val?.InnerText;
                            break;
                        case X.FontSize fontSize:
                            size = double.TryParse(
                                fontSize.Val?.InnerText,
                                System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out var parsed) ? parsed : null;
                            break;
                        case X.Bold b:
                            bold = ReadToggle(b.Val);
                            break;
                        case X.Italic i:
                            italic = ReadToggle(i.Val);
                            break;
                        case X.Underline u:
                            underline = u.Val?.InnerText != "none";
                            break;
                        case X.Strike st:
                            strike = ReadToggle(st.Val);
                            break;
                        case X.Color c:
                            color = TryResolveColor(c, colors);
                            break;
                        case X.VerticalTextAlignment v:
                            other |= v.Val?.InnerText is { Length: > 0 } align && align != "baseline";
                            break;
                        case X.Outline o:
                            other |= ReadToggle(o.Val);
                            break;
                        case X.Shadow sh:
                            other |= ReadToggle(sh.Val);
                            break;
                        case X.Condense cd:
                            other |= ReadToggle(cd.Val);
                            break;
                        case X.Extend ex:
                            other |= ReadToggle(ex.Val);
                            break;
                    }
                }

                return new RunFormat(name, size, bold, italic, underline, strike, color, other);
            }

            private static ArgbColor? TryResolveColor(X.Color color, ColorResolver colors)
            {
                try
                {
                    return colors.Resolve(color, ArgbColor.Black);
                }
                catch (FormatException)
                {
                    return null;
                }
            }

            /// <summary>
            /// <c>val</c> を省略した要素は有効。<c>0</c>・<c>false</c> は無効。それ以外(不正な値を含む)は有効とみなす。
            /// </summary>
            private static bool ReadToggle(DocumentFormat.OpenXml.BooleanValue? value) =>
                value?.InnerText is not ("0" or "false");

            public bool Differs(FontStyle font) =>
                HasOtherFormatting
                || (Name is not null && !string.Equals(Name, font.Name, StringComparison.OrdinalIgnoreCase))
                || (Size is { } size && Math.Abs(size - font.SizePt) > 0.01)
                || Bold != font.Bold
                || Italic != font.Italic
                || Underline != (font.Underline != UnderlineStyle.None)
                || Strike != font.Strike
                || (Color is { } color && !color.Equals(font.Color));
        }
    }
}
