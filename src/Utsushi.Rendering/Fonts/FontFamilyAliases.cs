using System;
using System.Collections.Generic;

namespace Utsushi.Rendering.Fonts
{
    /// <summary>
    /// 日本語のフォント名と英語のフォント名の対応(要件11.6)。
    /// </summary>
    /// <remarks>
    /// 日本語版のExcelはセル書式のフォント名を「ＭＳ Ｐゴシック」のような日本語名で保存する。
    /// 一方、実行環境(SkiaSharp)が返すファミリ名は英語名のことが多く、名前の完全一致だけで照合すると
    /// インストール済みのフォントを「見つからない」と誤判定する。帳票固有の情報ではなく、
    /// Windows・Office・主要な無償フォントに共通する名前の対応であるため、共通レイヤーに持つ。
    /// </remarks>
    internal static class FontFamilyAliases
    {
        private static readonly (string Japanese, string English)[] Pairs =
        {
            ("ＭＳ Ｐゴシック", "MS PGothic"),
            ("ＭＳ ゴシック", "MS Gothic"),
            ("ＭＳ Ｐ明朝", "MS PMincho"),
            ("ＭＳ 明朝", "MS Mincho"),
            ("ＭＳ ＵＩ Ｇｏｔｈｉｃ", "MS UI Gothic"),
            ("メイリオ", "Meiryo"),
            ("游ゴシック", "Yu Gothic"),
            ("游明朝", "Yu Mincho"),
            ("BIZ UDPゴシック", "BIZ UDPGothic"),
            ("BIZ UDゴシック", "BIZ UDGothic"),
            ("BIZ UDP明朝", "BIZ UDPMincho"),
            ("BIZ UD明朝", "BIZ UDMincho"),
            ("IPAゴシック", "IPAGothic"),
            ("IPA Pゴシック", "IPAPGothic"),
            ("IPA明朝", "IPAMincho"),
            ("IPA P明朝", "IPAPMincho"),
            ("IPAexゴシック", "IPAexGothic"),
            ("IPAex明朝", "IPAexMincho"),
            ("IPAmj明朝", "IPAmjMincho"),
        };

        private static readonly Dictionary<string, string> Map = BuildMap();

        /// <summary>指定の名前と、その別名(日本語名 ↔ 英語名)を返す。先頭は指定の名前そのもの。</summary>
        public static IEnumerable<string> Candidates(string familyName)
        {
            yield return familyName;
            if (Map.TryGetValue(Normalize(familyName), out var alias))
            {
                yield return alias;
            }
        }

        /// <summary>2つのファミリ名が同じフォントを指すかどうか(大文字小文字・別名を区別しない)。</summary>
        public static bool AreSame(string left, string right)
        {
            foreach (var candidate in Candidates(left))
            {
                if (string.Equals(Normalize(candidate), Normalize(right), StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, string> BuildMap()
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (japanese, english) in Pairs)
            {
                map[Normalize(japanese)] = english;
                map[Normalize(english)] = japanese;
            }

            return map;
        }

        /// <summary>全角英数・全角空白を半角にそろえる(「ＭＳ Ｐゴシック」と「MS Pゴシック」の表記揺れを吸収する)。</summary>
        private static string Normalize(string name)
        {
            var chars = name.Trim().ToCharArray();
            for (var i = 0; i < chars.Length; i++)
            {
                var c = chars[i];
                if (c >= '！' && c <= '～')
                {
                    chars[i] = (char)(c - 0xFEE0);
                }
                else if (c == '　')
                {
                    chars[i] = ' ';
                }
            }

            return new string(chars);
        }
    }
}
