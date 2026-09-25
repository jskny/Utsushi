using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SkiaSharp;

namespace Utsushi.Rendering.Fonts
{
    /// <summary>
    /// TrueType(<c>glyf</c> 形式)フォントから、帳票で実際に使った字形だけを含む小さなフォントを作る(要件11.5)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// SkiaSharp の PDF バックエンドはフォントを丸ごと埋め込む。IPAmj明朝(46MB)のような外字用フォントを
    /// 1文字使うだけで PDF が数十MBになるため、描画前にここで字形を絞り込んだフォントを作り、それを埋め込ませる。
    /// </para>
    /// <para>
    /// 字形番号は振り直す(<see cref="TrueTypeSubset.MapGlyph"/>)。<c>cmap</c> は使った文字と字形の対応だけで
    /// 作り直すため、PDF の文字列検索・コピー(ToUnicode)は元のフォントと同様に効く。字形の形・送り幅は
    /// 元の <c>glyf</c>/<c>hmtx</c> をそのまま写すため、見た目とレイアウトは変わらない。
    /// </para>
    /// <para>
    /// 次の場合はサブセットを作らない(<c>null</c> を返し、呼び出し側は元のフォントをそのまま使う)。
    /// CFF 形式などで <c>glyf</c> を持たない、埋め込み許可(OS/2 <c>fsType</c>)がサブセット化を禁止している、
    /// またはテーブルが壊れていて安全に読めない。
    /// </para>
    /// </remarks>
    internal static class TrueTypeSubsetter
    {
        private const ushort FsTypeRestrictedLicense = 0x0002;
        private const ushort FsTypeNoSubsetting = 0x0100;

        /// <summary>サブセットへ写すテーブル(字形の描画・PDF埋め込みに必要なもの)。<c>cmap</c>/<c>post</c> は作り直す。</summary>
        private static readonly string[] CopiedTables = { "OS/2", "name", "cvt ", "fpgm", "prep", "gasp" };

        /// <summary>
        /// サブセットを作る。
        /// </summary>
        /// <param name="typeface">元の書体。</param>
        /// <param name="glyphs">使った字形番号。</param>
        /// <param name="unicodeToGlyph">使った文字 → 元の字形番号(新しい <c>cmap</c> に載せる)。</param>
        public static TrueTypeSubset? TryCreate(
            SKTypeface typeface, IEnumerable<ushort> glyphs, IReadOnlyDictionary<int, ushort> unicodeToGlyph)
        {
            try
            {
                return Create(typeface, glyphs, unicodeToGlyph);
            }
            catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentException or InvalidDataException)
            {
                // テーブルが想定外の形をしている場合は、サブセット化をあきらめて元のフォントを埋め込む
                // (PDFが大きくなるだけで、描画結果は変わらない)。
                return null;
            }
        }

        private static TrueTypeSubset? Create(
            SKTypeface typeface, IEnumerable<ushort> glyphs, IReadOnlyDictionary<int, ushort> unicodeToGlyph)
        {
            var tables = typeface.GetTableTags().ToDictionary(TagName, tag => tag);
            if (!tables.ContainsKey("glyf") || !tables.ContainsKey("loca") || !tables.ContainsKey("head")
                || !tables.ContainsKey("hhea") || !tables.ContainsKey("hmtx") || !tables.ContainsKey("maxp"))
            {
                return null;
            }

            byte[] Table(string name) => typeface.GetTableData(tables[name]);

            if (tables.ContainsKey("OS/2"))
            {
                var os2 = Table("OS/2");
                var fsType = ReadUInt16(os2, 8);
                if ((fsType & (FsTypeRestrictedLicense | FsTypeNoSubsetting)) != 0)
                {
                    return null;
                }
            }

            var head = Table("head");
            var hhea = Table("hhea");
            var maxp = Table("maxp");
            var loca = Table("loca");
            var glyf = Table("glyf");
            var hmtx = Table("hmtx");

            var numGlyphs = ReadUInt16(maxp, 4);
            var longLoca = ReadInt16(head, 50) == 1;
            var numberOfHMetrics = ReadUInt16(hhea, 34);

            (int Start, int End) GlyphRange(int glyph)
            {
                if (glyph >= numGlyphs)
                {
                    return (0, 0);
                }

                var start = longLoca ? (int)ReadUInt32(loca, glyph * 4) : ReadUInt16(loca, glyph * 2) * 2;
                var end = longLoca ? (int)ReadUInt32(loca, (glyph + 1) * 4) : ReadUInt16(loca, (glyph + 1) * 2) * 2;
                if (start < 0 || end < start || end > glyf.Length)
                {
                    throw new InvalidDataException("loca が glyf の範囲外を指しています。");
                }

                return (start, end);
            }

            // 使った字形 + .notdef + 複合字形が参照する部品を集める。
            var kept = new SortedSet<ushort> { 0 };
            var pending = new Stack<ushort>(glyphs.Where(g => g < numGlyphs));
            while (pending.Count > 0)
            {
                var glyph = pending.Pop();
                if (!kept.Add(glyph) && glyph != 0)
                {
                    continue;
                }

                var (start, end) = GlyphRange(glyph);
                if (end - start < 10 || ReadInt16(glyf, start) >= 0)
                {
                    continue;
                }

                foreach (var (_, component) in EnumerateComponents(glyf, start, end))
                {
                    if (component < numGlyphs && !kept.Contains(component))
                    {
                        pending.Push(component);
                    }
                }
            }

            var order = kept.ToList();
            var oldToNew = new Dictionary<ushort, ushort>(order.Count);
            for (var i = 0; i < order.Count; i++)
            {
                oldToNew[order[i]] = (ushort)i;
            }

            // glyf / loca(長形式)を組み立てる。複合字形の部品番号は新しい番号へ書き換える。
            var newGlyf = new MemoryStream();
            var newLoca = new byte[(order.Count + 1) * 4];
            for (var i = 0; i < order.Count; i++)
            {
                WriteUInt32(newLoca, i * 4, (uint)newGlyf.Length);
                var (start, end) = GlyphRange(order[i]);
                if (end > start)
                {
                    var data = new byte[end - start];
                    Buffer.BlockCopy(glyf, start, data, 0, data.Length);
                    if (data.Length >= 10 && ReadInt16(data, 0) < 0)
                    {
                        foreach (var (offset, component) in EnumerateComponents(data, 0, data.Length))
                        {
                            WriteUInt16(data, offset, oldToNew.TryGetValue(component, out var mapped) ? mapped : (ushort)0);
                        }
                    }

                    newGlyf.Write(data, 0, data.Length);
                    while (newGlyf.Length % 4 != 0)
                    {
                        newGlyf.WriteByte(0);
                    }
                }
            }

            WriteUInt32(newLoca, order.Count * 4, (uint)newGlyf.Length);

            // hmtx: 全字形ぶんの (送り幅, 左側ベアリング) を持たせる。
            var newHmtx = new byte[order.Count * 4];
            for (var i = 0; i < order.Count; i++)
            {
                var glyph = order[i];
                ushort advance;
                short lsb;
                if (glyph < numberOfHMetrics)
                {
                    advance = ReadUInt16(hmtx, glyph * 4);
                    lsb = ReadInt16(hmtx, (glyph * 4) + 2);
                }
                else
                {
                    advance = ReadUInt16(hmtx, (numberOfHMetrics - 1) * 4);
                    var lsbOffset = (numberOfHMetrics * 4) + ((glyph - numberOfHMetrics) * 2);
                    lsb = lsbOffset + 2 <= hmtx.Length ? ReadInt16(hmtx, lsbOffset) : (short)0;
                }

                WriteUInt16(newHmtx, i * 4, advance);
                WriteUInt16(newHmtx, (i * 4) + 2, (ushort)lsb);
            }

            var newHead = (byte[])head.Clone();
            WriteUInt32(newHead, 8, 0); // checkSumAdjustment は最後に計算する
            WriteUInt16(newHead, 50, 1); // indexToLocFormat = long

            var newHhea = (byte[])hhea.Clone();
            WriteUInt16(newHhea, 34, (ushort)order.Count);

            var newMaxp = (byte[])maxp.Clone();
            WriteUInt16(newMaxp, 4, (ushort)order.Count);

            var mappedUnicode = unicodeToGlyph
                .Where(pair => oldToNew.ContainsKey(pair.Value))
                .ToDictionary(pair => pair.Key, pair => oldToNew[pair.Value]);

            var output = new SortedDictionary<string, byte[]>(StringComparer.Ordinal)
            {
                ["head"] = newHead,
                ["hhea"] = newHhea,
                ["maxp"] = newMaxp,
                ["hmtx"] = newHmtx,
                ["loca"] = newLoca,
                ["glyf"] = newGlyf.ToArray(),
                ["cmap"] = BuildCmap(mappedUnicode),
                ["post"] = BuildPost(tables.ContainsKey("post") ? Table("post") : null),
            };

            foreach (var name in CopiedTables)
            {
                if (tables.ContainsKey(name))
                {
                    output[name] = Table(name);
                }
            }

            return new TrueTypeSubset(AssembleSfnt(output), oldToNew);
        }

        /// <summary>複合字形の各部品について、(部品番号の位置, 部品番号) を列挙する。</summary>
        private static IEnumerable<(int Offset, ushort Glyph)> EnumerateComponents(byte[] data, int start, int end)
        {
            const ushort Arg1And2AreWords = 0x0001;
            const ushort WeHaveAScale = 0x0008;
            const ushort MoreComponents = 0x0020;
            const ushort WeHaveAnXAndYScale = 0x0040;
            const ushort WeHaveATwoByTwo = 0x0080;

            var offset = start + 10; // numberOfContours + バウンディングボックス
            ushort flags;
            do
            {
                if (offset + 4 > end)
                {
                    throw new InvalidDataException("複合字形の部品がglyfの範囲外です。");
                }

                flags = ReadUInt16(data, offset);
                var glyph = ReadUInt16(data, offset + 2);
                yield return (offset + 2, glyph);

                offset += 4;
                offset += (flags & Arg1And2AreWords) != 0 ? 4 : 2;
                if ((flags & WeHaveAScale) != 0)
                {
                    offset += 2;
                }
                else if ((flags & WeHaveAnXAndYScale) != 0)
                {
                    offset += 4;
                }
                else if ((flags & WeHaveATwoByTwo) != 0)
                {
                    offset += 8;
                }
            }
            while ((flags & MoreComponents) != 0);
        }

        /// <summary>format 4(BMP)と format 12(全面)の <c>cmap</c> を作る。</summary>
        private static byte[] BuildCmap(IReadOnlyDictionary<int, ushort> unicodeToGlyph)
        {
            var sorted = unicodeToGlyph.OrderBy(pair => pair.Key).ToList();

            // format 4: BMP の文字を1文字1セグメントで持つ(帳票1枚ぶんの文字数なら十分小さい)。
            var bmp = sorted.Where(pair => pair.Key <= 0xFFFF && pair.Key != 0xFFFF).ToList();
            var segCount = bmp.Count + 1; // 末尾の 0xFFFF 番兵
            var format4 = new byte[16 + (segCount * 8)];
            WriteUInt16(format4, 0, 4);
            WriteUInt16(format4, 2, (ushort)format4.Length);
            WriteUInt16(format4, 6, (ushort)(segCount * 2));
            var searchRange = 2 * (1 << (int)Math.Floor(Math.Log(segCount, 2)));
            WriteUInt16(format4, 8, (ushort)searchRange);
            WriteUInt16(format4, 10, (ushort)Math.Floor(Math.Log(searchRange / 2, 2)));
            WriteUInt16(format4, 12, (ushort)((segCount * 2) - searchRange));
            var endCodes = 14;
            var startCodes = endCodes + (segCount * 2) + 2;
            var idDeltas = startCodes + (segCount * 2);
            for (var i = 0; i < segCount; i++)
            {
                var code = i < bmp.Count ? (ushort)bmp[i].Key : (ushort)0xFFFF;
                var glyph = i < bmp.Count ? bmp[i].Value : (ushort)0;
                WriteUInt16(format4, endCodes + (i * 2), code);
                WriteUInt16(format4, startCodes + (i * 2), code);
                WriteUInt16(format4, idDeltas + (i * 2), i < bmp.Count ? (ushort)((glyph - code) & 0xFFFF) : (ushort)1);
                // idRangeOffset はすべて 0(idDelta で直接対応させる)
            }

            var format12 = new byte[16 + (sorted.Count * 12)];
            WriteUInt16(format12, 0, 12);
            WriteUInt32(format12, 4, (uint)format12.Length);
            WriteUInt32(format12, 12, (uint)sorted.Count);
            for (var i = 0; i < sorted.Count; i++)
            {
                var o = 16 + (i * 12);
                WriteUInt32(format12, o, (uint)sorted[i].Key);
                WriteUInt32(format12, o + 4, (uint)sorted[i].Key);
                WriteUInt32(format12, o + 8, sorted[i].Value);
            }

            var header = new byte[4 + (2 * 8)];
            WriteUInt16(header, 2, 2);
            WriteUInt16(header, 4, 3); // Windows
            WriteUInt16(header, 6, 1); // Unicode BMP
            WriteUInt32(header, 8, (uint)header.Length);
            WriteUInt16(header, 12, 3); // Windows
            WriteUInt16(header, 14, 10); // Unicode full
            WriteUInt32(header, 16, (uint)(header.Length + format4.Length));

            return header.Concat(format4).Concat(format12).ToArray();
        }

        /// <summary>字形名を持たない <c>post</c>(format 3)を作る。元の斜体角・下線位置などは引き継ぐ。</summary>
        private static byte[] BuildPost(byte[]? original)
        {
            var post = new byte[32];
            if (original is { Length: >= 32 })
            {
                Buffer.BlockCopy(original, 0, post, 0, 32);
            }

            WriteUInt32(post, 0, 0x00030000);
            return post;
        }

        private static byte[] AssembleSfnt(SortedDictionary<string, byte[]> tables)
        {
            var numTables = tables.Count;
            var entrySelector = (int)Math.Floor(Math.Log(numTables, 2));
            var searchRange = (1 << entrySelector) * 16;

            var headerLength = 12 + (numTables * 16);
            var output = new MemoryStream();
            var header = new byte[headerLength];
            WriteUInt32(header, 0, 0x00010000);
            WriteUInt16(header, 4, (ushort)numTables);
            WriteUInt16(header, 6, (ushort)searchRange);
            WriteUInt16(header, 8, (ushort)entrySelector);
            WriteUInt16(header, 10, (ushort)((numTables * 16) - searchRange));
            output.Write(header, 0, header.Length);

            var headOffset = 0;
            var index = 0;
            foreach (var (name, data) in tables)
            {
                var offset = (int)output.Length;
                if (name == "head")
                {
                    headOffset = offset;
                }

                var record = 12 + (index * 16);
                for (var i = 0; i < 4; i++)
                {
                    header[record + i] = (byte)name[i];
                }

                WriteUInt32(header, record + 4, Checksum(data));
                WriteUInt32(header, record + 8, (uint)offset);
                WriteUInt32(header, record + 12, (uint)data.Length);

                output.Write(data, 0, data.Length);
                while (output.Length % 4 != 0)
                {
                    output.WriteByte(0);
                }

                index++;
            }

            var bytes = output.ToArray();
            Buffer.BlockCopy(header, 0, bytes, 0, header.Length);
            WriteUInt32(bytes, headOffset + 8, unchecked(0xB1B0AFBA - Checksum(bytes)));
            return bytes;
        }

        private static uint Checksum(byte[] data)
        {
            uint sum = 0;
            for (var i = 0; i < data.Length; i += 4)
            {
                uint value = 0;
                for (var j = 0; j < 4; j++)
                {
                    value = (value << 8) | (i + j < data.Length ? data[i + j] : (byte)0);
                }

                sum = unchecked(sum + value);
            }

            return sum;
        }

        private static string TagName(uint tag) => new(new[]
        {
            (char)(tag >> 24), (char)((tag >> 16) & 0xFF), (char)((tag >> 8) & 0xFF), (char)(tag & 0xFF),
        });

        internal static ushort ReadUInt16(byte[] data, int offset) => (ushort)((data[offset] << 8) | data[offset + 1]);

        internal static short ReadInt16(byte[] data, int offset) => (short)ReadUInt16(data, offset);

        internal static uint ReadUInt32(byte[] data, int offset) =>
            ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];

        private static void WriteUInt16(byte[] data, int offset, ushort value)
        {
            data[offset] = (byte)(value >> 8);
            data[offset + 1] = (byte)value;
        }

        private static void WriteUInt32(byte[] data, int offset, uint value)
        {
            data[offset] = (byte)(value >> 24);
            data[offset + 1] = (byte)(value >> 16);
            data[offset + 2] = (byte)(value >> 8);
            data[offset + 3] = (byte)value;
        }
    }

    /// <summary><see cref="TrueTypeSubsetter"/> が作ったサブセットフォント。</summary>
    internal sealed class TrueTypeSubset
    {
        private readonly IReadOnlyDictionary<ushort, ushort> _oldToNew;

        public TrueTypeSubset(byte[] fontData, IReadOnlyDictionary<ushort, ushort> oldToNew)
        {
            FontData = fontData;
            _oldToNew = oldToNew;
        }

        /// <summary>サブセットフォントのバイト列(sfnt)。</summary>
        public byte[] FontData { get; }

        /// <summary>元の字形番号を、サブセット内の字形番号に変換する。含まれない字形は .notdef(0)。</summary>
        public ushort MapGlyph(ushort original) => _oldToNew.TryGetValue(original, out var mapped) ? mapped : (ushort)0;
    }
}
