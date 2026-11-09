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

        private static IReadOnlyList<RunFormat>? ReadRuns(IEnumerable<X.Run> runs, ColorResolver colors)
        {
            List<RunFormat>? formats = null;
            foreach (var run in runs)
            {
                if (run.RunProperties is { } properties)
                {
                    (formats ??= new List<RunFormat>()).Add(RunFormat.Read(properties, colors));
                }
            }

            return formats;
        }

        private static bool Differs(IReadOnlyList<RunFormat> runs, FontStyle font)
        {
            foreach (var run in runs)
            {
                if (run.Differs(font))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>1つのランの書式(指定のあるものだけ)。</summary>
        private sealed record RunFormat(
            string? Name, double? Size, bool? Bold, bool? Italic, bool? Underline, bool? Strike, ArgbColor? Color, bool HasOtherFormatting)
        {
            public static RunFormat Read(X.RunProperties properties, ColorResolver colors)
            {
                string? name = null;
                double? size = null;
                bool? bold = null, italic = null, underline = null, strike = null;
                ArgbColor? color = null;
                var other = false;

                foreach (var element in properties.ChildElements)
                {
                    switch (element)
                    {
                        case X.RunFont runFont:
                            name = runFont.Val?.Value;
                            break;
                        case X.FontSize fontSize:
                            size = fontSize.Val?.Value;
                            break;
                        case X.Bold b:
                            bold = b.Val?.Value ?? true;
                            break;
                        case X.Italic i:
                            italic = i.Val?.Value ?? true;
                            break;
                        case X.Underline u:
                            underline = u.Val?.Value != X.UnderlineValues.None;
                            break;
                        case X.Strike s:
                            strike = s.Val?.Value ?? true;
                            break;
                        case X.Color c:
                            color = colors.Resolve(c, ArgbColor.Black);
                            break;
                        case X.VerticalTextAlignment v:
                            other |= v.Val?.Value is { } align && align != X.VerticalAlignmentRunValues.Baseline;
                            break;
                        case X.Outline o:
                            other |= o.Val?.Value ?? true;
                            break;
                        case X.Shadow sh:
                            other |= sh.Val?.Value ?? true;
                            break;
                        case X.Condense cd:
                            other |= cd.Val?.Value ?? true;
                            break;
                        case X.Extend ex:
                            other |= ex.Val?.Value ?? true;
                            break;
                    }
                }

                return new RunFormat(name, size, bold, italic, underline, strike, color, other);
            }

            public bool Differs(FontStyle font) =>
                HasOtherFormatting
                || (Name is not null && !string.Equals(Name, font.Name, StringComparison.OrdinalIgnoreCase))
                || (Size is { } size && Math.Abs(size - font.SizePt) > 0.01)
                || (Bold is { } bold && bold != font.Bold)
                || (Italic is { } italic && italic != font.Italic)
                || (Underline is { } underline && underline != (font.Underline != UnderlineStyle.None))
                || (Strike is { } strike && strike != font.Strike)
                || (Color is { } color && !color.Equals(font.Color));
        }
    }
}
