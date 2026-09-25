using System;
using System.Collections.Generic;
using SkiaSharp;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Utsushi.Rendering.Fonts;

namespace Utsushi.Rendering
{
    /// <summary>
    /// 1回のPDF出力で使う、字形の並びとサブセットフォントの一式。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 描画の前に全ページの文字列(セル・ヘッダー/フッター・図形内テキスト)を <see cref="GlyphShaper"/> で字形の並びに変換し、
    /// 字形の欠落をページの描画を始める前に検出する(要件5.5)。
    /// </para>
    /// <para>
    /// フォント埋め込みの場合は、書体ごとに使った字形を集めてサブセットフォントを作る(要件11.5)。
    /// サブセットを作れない書体(CFF形式、埋め込み許可がサブセット化を禁止)は元の書体をそのまま埋め込む。
    /// 出力ごとに作るため、複数スレッドから同じレンダラーを使っても互いに干渉しない。
    /// </para>
    /// </remarks>
    internal sealed class RenderContext : IDisposable
    {
        private readonly Dictionary<(FontStyle Font, string Text), ShapedText> _shaped;
        private readonly Dictionary<SKTypeface, SubsetGroup> _subsets;

        private RenderContext(
            string reportCode,
            string sheetName,
            Dictionary<(FontStyle, string), ShapedText> shaped,
            Dictionary<SKTypeface, SubsetGroup> subsets)
        {
            ReportCode = reportCode;
            SheetName = sheetName;
            _shaped = shaped;
            _subsets = subsets;
        }

        public string ReportCode { get; }

        public string SheetName { get; }

        /// <summary>全ページの文字列を字形の並びに変換し、必要ならサブセットフォントを作る。</summary>
        /// <exception cref="Utsushi.Core.Exceptions.MissingGlyphException">
        /// 字形の無い文字があり、<see cref="PdfRenderOptions.MissingGlyphs"/> が <see cref="MissingGlyphPolicy.Error"/> の場合。
        /// </exception>
        public static RenderContext Prepare(PagedLayout layout, GlyphShaper shaper, PdfRenderOptions options)
        {
            var shaped = new Dictionary<(FontStyle, string), ShapedText>();
            foreach (var page in layout.Pages)
            {
                foreach (var (font, text) in EnumerateTexts(page.Commands))
                {
                    if (shaped.ContainsKey((font, text)))
                    {
                        continue;
                    }

                    var result = shaper.Shape(font, text);
                    if (result.Missing is not null && options.MissingGlyphs == MissingGlyphPolicy.Error)
                    {
                        throw SkiaPdfRenderer.CreateMissingGlyphException(result, text, layout.ReportCode, layout.SheetName);
                    }

                    shaped[(font, text)] = result;
                }
            }

            var subsets = new Dictionary<SKTypeface, SubsetGroup>();
            if (options.TextRendering == PdfTextRendering.EmbedFont)
            {
                try
                {
                    BuildSubsets(shaped.Values, subsets);
                }
                catch
                {
                    DisposeSubsets(subsets);
                    throw;
                }
            }

            return new RenderContext(layout.ReportCode, layout.SheetName, shaped, subsets);
        }

        /// <summary><see cref="Prepare"/> で変換済みの字形の並びを返す。</summary>
        public ShapedText GetShaped(FontStyle font, string text) => _shaped[(font, text)];

        /// <summary>
        /// 埋め込みに使う書体と字形番号の並びを返す。サブセットがあればサブセット側の番号に振り直す。
        /// 異体字のために同じ書体のサブセットが複数ある場合は、サブセットごとに区切った並びになる。
        /// </summary>
        public IEnumerable<(SKTypeface Typeface, ushort[] Glyphs)> MapForEmbedding(GlyphRun run)
        {
            if (!_subsets.TryGetValue(run.Face.Typeface, out var group))
            {
                yield return (run.Face.Typeface, run.Glyphs);
                yield break;
            }

            var start = 0;
            while (start < run.Glyphs.Length)
            {
                var bucket = group.BucketOf(run.Glyphs[start]);
                var end = start + 1;
                while (end < run.Glyphs.Length && group.BucketOf(run.Glyphs[end]) == bucket)
                {
                    end++;
                }

                var (typeface, subset) = group.Subsets[bucket];
                var glyphs = new ushort[end - start];
                for (var i = 0; i < glyphs.Length; i++)
                {
                    glyphs[i] = subset.MapGlyph(run.Glyphs[start + i]);
                }

                yield return (typeface, glyphs);
                start = end;
            }
        }

        public void Dispose() => DisposeSubsets(_subsets);

        /// <summary>
        /// 書体ごとに使った字形を集め、サブセットフォントを作る。
        /// </summary>
        /// <remarks>
        /// <c>cmap</c> は1つの文字に1つの字形しか対応づけられない。同じ基底文字の通常の字形と異体字の字形を
        /// 同じ文書で使う場合(例: 「葛」と「葛」+IVS)は、後から出てきた字形を別のサブセット(バケット)に入れ、
        /// どちらの字形も基底文字として文字列抽出できるようにする。対応づけの無い字形があると、SkiaSharp は
        /// ToUnicode を U+0000 にし、PDFビューアによってはそこで文字列の抽出が途切れる(実測)。
        /// </remarks>
        private static void BuildSubsets(IEnumerable<ShapedText> shapedTexts, Dictionary<SKTypeface, SubsetGroup> subsets)
        {
            var usage = new Dictionary<SKTypeface, SubsetUsage>();
            foreach (var shaped in shapedTexts)
            {
                foreach (var run in shaped.Runs)
                {
                    if (!usage.TryGetValue(run.Face.Typeface, out var entry))
                    {
                        entry = new SubsetUsage();
                        usage[run.Face.Typeface] = entry;
                    }

                    foreach (var (codePoint, glyph) in run.CharacterMap)
                    {
                        entry.Add(glyph, codePoint);
                    }

                    foreach (var glyph in run.Glyphs)
                    {
                        entry.Add(glyph, null);
                    }
                }
            }

            foreach (var (typeface, entry) in usage)
            {
                var created = new List<(SKTypeface, TrueTypeSubset)>();
                foreach (var (glyphs, characterMap) in entry.Buckets)
                {
                    var subset = TrueTypeSubsetter.TryCreate(typeface, glyphs, characterMap);
                    if (subset is null)
                    {
                        break;
                    }

                    using var data = SKData.CreateCopy(subset.FontData);
                    var subsetTypeface = SKTypeface.FromData(data);
                    if (subsetTypeface is null)
                    {
                        break;
                    }

                    created.Add((subsetTypeface, subset));
                }

                if (created.Count == entry.Buckets.Count)
                {
                    subsets[typeface] = new SubsetGroup(entry.GlyphBuckets, created);
                    continue;
                }

                // サブセットを作れない書体(CFF形式・埋め込み許可の制限など)は元の書体を埋め込む。
                foreach (var (subsetTypeface, _) in created)
                {
                    subsetTypeface.Dispose();
                }
            }
        }

        private static void DisposeSubsets(Dictionary<SKTypeface, SubsetGroup> subsets)
        {
            foreach (var group in subsets.Values)
            {
                foreach (var (typeface, _) in group.Subsets)
                {
                    typeface.Dispose();
                }
            }

            subsets.Clear();
        }

        /// <summary>1つの書体について、字形をどのサブセット(バケット)に入れるかを決める。</summary>
        private sealed class SubsetUsage
        {
            public List<(HashSet<ushort> Glyphs, Dictionary<int, ushort> CharacterMap)> Buckets { get; } = new();

            public Dictionary<ushort, int> GlyphBuckets { get; } = new();

            public void Add(ushort glyph, int? codePoint)
            {
                if (GlyphBuckets.TryGetValue(glyph, out var assigned))
                {
                    if (codePoint is { } cp && !Buckets[assigned].CharacterMap.ContainsKey(cp))
                    {
                        Buckets[assigned].CharacterMap[cp] = glyph;
                    }

                    return;
                }

                var bucket = 0;
                if (codePoint is { } value)
                {
                    while (bucket < Buckets.Count
                        && Buckets[bucket].CharacterMap.TryGetValue(value, out var existing) && existing != glyph)
                    {
                        bucket++;
                    }
                }

                if (bucket == Buckets.Count)
                {
                    Buckets.Add((new HashSet<ushort>(), new Dictionary<int, ushort>()));
                }

                Buckets[bucket].Glyphs.Add(glyph);
                if (codePoint is { } mapped)
                {
                    Buckets[bucket].CharacterMap[mapped] = glyph;
                }

                GlyphBuckets[glyph] = bucket;
            }
        }

        /// <summary>1つの書体から作ったサブセットの一式。</summary>
        private sealed class SubsetGroup
        {
            private readonly Dictionary<ushort, int> _glyphBuckets;

            public SubsetGroup(Dictionary<ushort, int> glyphBuckets, List<(SKTypeface, TrueTypeSubset)> subsets)
            {
                _glyphBuckets = glyphBuckets;
                Subsets = subsets;
            }

            public List<(SKTypeface Typeface, TrueTypeSubset Subset)> Subsets { get; }

            public int BucketOf(ushort glyph) => _glyphBuckets.TryGetValue(glyph, out var bucket) ? bucket : 0;
        }

        private static IEnumerable<(FontStyle Font, string Text)> EnumerateTexts(IEnumerable<DrawCommand> commands)
        {
            foreach (var command in commands)
            {
                switch (command)
                {
                    case TextCommand text:
                        yield return (text.Font, text.Text);
                        break;
                    case ShapeCommand shape:
                        foreach (var line in shape.TextLines)
                        {
                            yield return (line.Font, line.Text);
                        }

                        break;
                    case GroupCommand group:
                        foreach (var child in EnumerateTexts(group.Children))
                        {
                            yield return child;
                        }

                        break;
                }
            }
        }
    }
}
