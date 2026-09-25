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
            try
            {
                var cmap = Array.IndexOf(typeface.GetTableTags(), CmapTag) >= 0 ? typeface.GetTableData(CmapTag) : null;
                return cmap is null ? Empty : Parse(cmap);
            }
            catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentException or InvalidDataException)
            {
                // 壊れた表は「異体字を登録していない」ものとして扱う(該当文字は要件5.5のエラーになる)。
                return Empty;
            }
        }

        private static VariationSequenceTable Parse(byte[] cmap)
        {
            var selectors = new Dictionary<int, (List<(int, int)>, Dictionary<int, ushort>)>();
            var numTables = TrueTypeSubsetter.ReadUInt16(cmap, 2);
            for (var i = 0; i < numTables; i++)
            {
                var subtable = (int)TrueTypeSubsetter.ReadUInt32(cmap, 4 + (i * 8) + 4);
                if (subtable + 10 > cmap.Length || TrueTypeSubsetter.ReadUInt16(cmap, subtable) != 14)
                {
                    continue;
                }

                var count = TrueTypeSubsetter.ReadUInt32(cmap, subtable + 6);
                for (var r = 0; r < count; r++)
                {
                    var record = subtable + 10 + (r * 11);
                    var selector = ReadUInt24(cmap, record);
                    var defaultOffset = (int)TrueTypeSubsetter.ReadUInt32(cmap, record + 3);
                    var nonDefaultOffset = (int)TrueTypeSubsetter.ReadUInt32(cmap, record + 7);

                    var defaults = new List<(int, int)>();
                    if (defaultOffset != 0)
                    {
                        var at = subtable + defaultOffset;
                        var ranges = TrueTypeSubsetter.ReadUInt32(cmap, at);
                        for (var k = 0; k < ranges; k++)
                        {
                            var start = ReadUInt24(cmap, at + 4 + (k * 4));
                            defaults.Add((start, start + cmap[at + 4 + (k * 4) + 3]));
                        }
                    }

                    var nonDefaults = new Dictionary<int, ushort>();
                    if (nonDefaultOffset != 0)
                    {
                        var at = subtable + nonDefaultOffset;
                        var mappings = TrueTypeSubsetter.ReadUInt32(cmap, at);
                        for (var k = 0; k < mappings; k++)
                        {
                            var o = at + 4 + (k * 5);
                            nonDefaults[ReadUInt24(cmap, o)] = TrueTypeSubsetter.ReadUInt16(cmap, o + 3);
                        }
                    }

                    selectors[selector] = (defaults, nonDefaults);
                }

                break;
            }

            return new VariationSequenceTable(selectors);
        }

        private static int ReadUInt24(byte[] data, int offset) => (data[offset] << 16) | (data[offset + 1] << 8) | data[offset + 2];
    }
}
