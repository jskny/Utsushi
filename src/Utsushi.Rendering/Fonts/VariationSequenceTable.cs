using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using SkiaSharp;

namespace Utsushi.Rendering.Fonts
{
    /// <summary>
    /// フォントの異体字シーケンス表(OpenType <c>cmap</c> format 14)(要件11.3)。
    /// </summary>
    /// <remarks>
    /// SkiaSharp の文字→字形の変換(<c>SKFont.GetGlyph</c>)は1文字ずつで、異体字セレクタを解釈しない。
    /// 「基底文字 + 異体字セレクタ」の組から字形を選ぶため、ここで format 14 を直接読む。
    /// </remarks>
    internal sealed class VariationSequenceTable
    {
        private static readonly ConditionalWeakTable<SKTypeface, VariationSequenceTable> Cache = new();

        private static readonly VariationSequenceTable Empty = new(
            new Dictionary<int, (List<(int Start, int End)> Default, Dictionary<int, ushort> NonDefault)>());

        private readonly Dictionary<int, (List<(int Start, int End)> Default, Dictionary<int, ushort> NonDefault)> _selectors;

        private VariationSequenceTable(
            Dictionary<int, (List<(int Start, int End)> Default, Dictionary<int, ushort> NonDefault)> selectors)
        {
            _selectors = selectors;
        }

        /// <summary>異体字セレクタかどうか(SVS: U+FE00〜FE0F、IVS: U+E0100〜E01EF)。</summary>
        public static bool IsVariationSelector(int codePoint) =>
            (codePoint >= 0xFE00 && codePoint <= 0xFE0F) || (codePoint >= 0xE0100 && codePoint <= 0xE01EF);

        /// <summary>書体の異体字シーケンス表を返す(書体ごとに1回だけ読む)。</summary>
        public static VariationSequenceTable For(SKTypeface typeface) =>
            Cache.GetValue(typeface, Load);

        /// <summary>
        /// 「基底文字 + 異体字セレクタ」の字形を探す。
        /// </summary>
        /// <param name="baseCodePoint">基底文字。</param>
        /// <param name="selector">異体字セレクタ。</param>
        /// <param name="isDefault">
        /// true なら、この組は基底文字の通常の字形で表す(Default UVS)。呼び出し側は <c>cmap</c> で基底文字を引く。
        /// </param>
        /// <param name="glyph">通常の字形と異なる字形(Non-Default UVS)の字形番号。</param>
        /// <returns>フォントがこの組を登録している場合 true。</returns>
        public bool TryLookup(int baseCodePoint, int selector, out bool isDefault, out ushort glyph)
        {
            isDefault = false;
            glyph = 0;
            if (!_selectors.TryGetValue(selector, out var entry))
            {
                return false;
            }

            if (entry.NonDefault.TryGetValue(baseCodePoint, out glyph))
            {
                return true;
            }

            foreach (var (start, end) in entry.Default)
            {
                if (baseCodePoint >= start && baseCodePoint <= end)
                {
                    isDefault = true;
                    return true;
                }
            }

            return false;
        }

        private static VariationSequenceTable Load(SKTypeface typeface)
        {
            const uint CmapTag = 0x636D6170;
            const uint MaxpTag = 0x6D617870;
            try
            {
                // SkiaSharp の GetTableData は、表を読めない場合に基底の Exception を投げるため TryGetTableData を使う。
                if (!typeface.TryGetTableData(CmapTag, out var cmap) || cmap is null)
                {
                    return Empty;
                }

                var numGlyphs = typeface.TryGetTableData(MaxpTag, out var maxp) && maxp is { Length: >= 6 }
                    ? TrueTypeSubsetter.ReadUInt16(maxp, 4)
                    : ushort.MaxValue;
                return Parse(cmap, numGlyphs);
            }
            catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentException or InvalidDataException)
            {
                // 壊れた表は「異体字を登録していない」ものとして扱う(該当文字は要件5.5のエラーになる)。
                return Empty;
            }
        }

        /// <summary>
        /// format 14 を読む。
        /// </summary>
        /// <remarks>
        /// 件数は、表の残りのバイト数に収まる範囲かを読む前に確かめる。また、複数のレコードが同じ
        /// Default/Non-Default 表を指すことを許しつつ、同じ表は一度だけ読んで共有する。さらに読み取る項目の
        /// 総数を表の長さに比例する上限で打ち切る。これらが無いと、小さな表でも「レコード数 × 共有された表の
        /// 件数」に比例する処理量・メモリを使わせることができる(security-reviewerの実測で300KBの表から約6GB)。
        /// 字形番号は字形数(<c>maxp</c>)の範囲外なら登録が無いものとして扱う(範囲外の番号を「見つかった」と
        /// みなすと、描画で .notdef になり字形欠落の検出をすり抜けるため)。
        /// </remarks>
        private static VariationSequenceTable Parse(byte[] cmap, ushort numGlyphs)
        {
            var selectors = new Dictionary<int, (List<(int, int)>, Dictionary<int, ushort>)>();
            var defaultTables = new Dictionary<int, List<(int, int)>>();
            var nonDefaultTables = new Dictionary<int, Dictionary<int, ushort>>();
            var budget = cmap.Length; // 読み取る項目の総数の上限(1項目は最低4バイトのため、正しい表なら超えない)

            var numTables = TrueTypeSubsetter.ReadUInt16(cmap, 2);
            for (var i = 0; i < numTables; i++)
            {
                var subtable = (int)TrueTypeSubsetter.ReadUInt32(cmap, 4 + (i * 8) + 4);
                if (subtable < 0 || subtable + 10 > cmap.Length || TrueTypeSubsetter.ReadUInt16(cmap, subtable) != 14)
                {
                    continue;
                }

                var count = TrueTypeSubsetter.ReadUInt32(cmap, subtable + 6);
                if (count > (uint)((cmap.Length - subtable - 10) / 11))
                {
                    throw new InvalidDataException("cmap format 14 のレコード数が表の長さを超えています。");
                }

                for (var r = 0; r < count; r++)
                {
                    var record = subtable + 10 + (r * 11);
                    var selector = ReadUInt24(cmap, record);
                    var defaultOffset = (int)TrueTypeSubsetter.ReadUInt32(cmap, record + 3);
                    var nonDefaultOffset = (int)TrueTypeSubsetter.ReadUInt32(cmap, record + 7);

                    var defaults = defaultOffset == 0
                        ? new List<(int, int)>()
                        : GetOrRead(defaultTables, subtable + defaultOffset, at => ReadDefaultTable(cmap, at, ref budget));
                    var nonDefaults = nonDefaultOffset == 0
                        ? new Dictionary<int, ushort>()
                        : GetOrRead(nonDefaultTables, subtable + nonDefaultOffset, at => ReadNonDefaultTable(cmap, at, numGlyphs, ref budget));

                    selectors[selector] = (defaults, nonDefaults);
                }

                break;
            }

            return new VariationSequenceTable(selectors);
        }

        private delegate T TableReader<out T>(int offset);

        private static T GetOrRead<T>(Dictionary<int, T> cache, int offset, TableReader<T> read)
        {
            if (!cache.TryGetValue(offset, out var table))
            {
                table = read(offset);
                cache[offset] = table;
            }

            return table;
        }

        private static List<(int, int)> ReadDefaultTable(byte[] cmap, int at, ref int budget)
        {
            if (at < 0 || at + 4 > cmap.Length)
            {
                throw new InvalidDataException("cmap format 14 の Default UVS 表が範囲外です。");
            }

            var ranges = TrueTypeSubsetter.ReadUInt32(cmap, at);
            if (ranges > (uint)((cmap.Length - at - 4) / 4) || (budget -= (int)ranges) < 0)
            {
                throw new InvalidDataException("cmap format 14 の Default UVS の件数が不正です。");
            }

            var list = new List<(int, int)>((int)ranges);
            for (var k = 0; k < ranges; k++)
            {
                var start = ReadUInt24(cmap, at + 4 + (k * 4));
                list.Add((start, start + cmap[at + 4 + (k * 4) + 3]));
            }

            return list;
        }

        private static Dictionary<int, ushort> ReadNonDefaultTable(byte[] cmap, int at, ushort numGlyphs, ref int budget)
        {
            if (at < 0 || at + 4 > cmap.Length)
            {
                throw new InvalidDataException("cmap format 14 の Non-Default UVS 表が範囲外です。");
            }

            var mappings = TrueTypeSubsetter.ReadUInt32(cmap, at);
            if (mappings > (uint)((cmap.Length - at - 4) / 5) || (budget -= (int)mappings) < 0)
            {
                throw new InvalidDataException("cmap format 14 の Non-Default UVS の件数が不正です。");
            }

            var map = new Dictionary<int, ushort>((int)mappings);
            for (var k = 0; k < mappings; k++)
            {
                var o = at + 4 + (k * 5);
                var glyph = TrueTypeSubsetter.ReadUInt16(cmap, o + 3);
                if (glyph < numGlyphs)
                {
                    map[ReadUInt24(cmap, o)] = glyph;
                }
            }

            return map;
        }

        private static int ReadUInt24(byte[] data, int offset) => (data[offset] << 16) | (data[offset + 1] << 8) | data[offset + 2];
    }
}
