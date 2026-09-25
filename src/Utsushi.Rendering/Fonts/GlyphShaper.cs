using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using SkiaSharp;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering.Fonts
{
    /// <summary>
    /// 文字列を「どの書体のどの字形で描くか」の並び(<see cref="GlyphRun"/>)に変換する(要件11.2〜11.4)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 計測(<see cref="SkiaFontMetricsProvider"/>)と描画(<see cref="SkiaPdfRenderer"/>)の両方がこの結果を使うため、
    /// 外字用の代替フォントで描いた文字の送り幅もレイアウト計算に正しく反映される。
    /// </para>
    /// <para>
    /// 書記素クラスタ(利用者が1文字と認識する単位)ごとに、セルのフォント → 外字用の代替フォントの順で字形を探す。
    /// 異体字セレクタ付きの文字は <c>cmap</c> format 14(<see cref="VariationSequenceTable"/>)で字形を選ぶ。
    /// 仮名 + 合成用濁点・半濁点は合成済みの仮名に置き換える。それ以外の正規化(CJK互換漢字の統合漢字への
    /// 置き換えなど)は、人名の字形が変わってしまうため行わない。
    /// </para>
    /// </remarks>
    internal sealed class GlyphShaper
    {
        private readonly FontResolver _fontResolver;

        public GlyphShaper(FontResolver fontResolver)
        {
            _fontResolver = fontResolver ?? throw new ArgumentNullException(nameof(fontResolver));
        }

        /// <summary>文字列を字形の並びに変換する。</summary>
        public ShapedText Shape(FontStyle font, string text)
        {
            var primary = _fontResolver.Resolve(font);
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedText(primary, Array.Empty<GlyphRun>(), null);
            }

            if (TryShapeWithPrimaryOnly(primary, text) is { } simple)
            {
                return simple;
            }

            var chain = new List<ResolvedTypeface> { primary };
            foreach (var fallback in _fontResolver.ResolveGlyphFallbacks(font.Bold, font.Italic))
            {
                if (!ReferenceEquals(fallback.Typeface, primary.Typeface))
                {
                    chain.Add(fallback);
                }
            }

            var builder = new RunBuilder();
            MissingCharacter? missing = null;

            var elements = StringInfo.GetTextElementEnumerator(text);
            while (elements.MoveNext())
            {
                var codePoints = ToCodePoints(ComposeKanaVoicedMark(elements.GetTextElement()));
                var index = 0;

                // 基底文字 + 異体字セレクタ(要件11.3)
                if (codePoints.Count >= 2 && VariationSelectorAt(codePoints, 1))
                {
                    if (TryResolveVariation(chain, codePoints[0], codePoints[1], out var face, out var glyph, out _))
                    {
                        // 異体字の字形も基底文字に対応づける(PDFから文字列を抽出したときに基底文字として取り出せるように。
                        // 同じ基底文字の通常の字形と衝突する場合は、埋め込み時に別のサブセットへ分ける。RenderContext 参照)。
                        builder.Add(face, glyph, codePoints[0]);
                    }
                    else
                    {
                        missing ??= new MissingCharacter(codePoints[1], codePoints[0]);

                        // 描画を続ける設定の場合は、異体字の指定を諦めて基底文字だけを描く(豆腐を出さない)。
                        AddCodePoint(builder, chain, codePoints[0], ref missing);
                    }

                    index = 2;
                }

                for (; index < codePoints.Count; index++)
                {
                    var codePoint = codePoints[index];
                    if (Rune.IsControl(new Rune(codePoint)) || VariationSequenceTable.IsVariationSelector(codePoint))
                    {
                        continue;
                    }

                    AddCodePoint(builder, chain, codePoint, ref missing);
                }
            }

            return new ShapedText(primary, builder.Build(), missing);
        }

        /// <summary>計測・描画用の <see cref="SKFont"/> を作る。斜体の合成もここで与える。</summary>
        public static SKFont CreateFont(ResolvedTypeface face, double sizePt, SKTypeface? typefaceOverride = null)
        {
            var font = new SKFont(typefaceOverride ?? face.Typeface, (float)sizePt)
            {
                Subpixel = true,
                Edging = SKFontEdging.Antialias,
            };

            if (face.SynthesizeItalic)
            {
                font.SkewX = SkiaFontMetricsProvider.ItalicSkew;
            }

            return font;
        }

        /// <summary>
        /// 大半の文字列は、セルのフォントだけで全文字の字形がそろう。その場合は1つの並びにまとめて返す(計測が頻繁なため)。
        /// </summary>
        private static ShapedText? TryShapeWithPrimaryOnly(ResolvedTypeface primary, string text)
        {
            var codePoints = new List<int>(text.Length);
            foreach (var rune in text.EnumerateRunes())
            {
                var value = rune.Value;
                if (Rune.IsControl(rune) || VariationSequenceTable.IsVariationSelector(value)
                    || value == 0x3099 || value == 0x309A || value == 0xFFFD)
                {
                    return null;
                }

                codePoints.Add(value);
            }

            var glyphs = new ushort[codePoints.Count];
            for (var i = 0; i < codePoints.Count; i++)
            {
                glyphs[i] = primary.Typeface.GetGlyph(codePoints[i]);
                if (glyphs[i] == 0)
                {
                    return null;
                }
            }

            var cmap = new List<(int, ushort)>(codePoints.Count);
            for (var i = 0; i < codePoints.Count; i++)
            {
                cmap.Add((codePoints[i], glyphs[i]));
            }

            return new ShapedText(primary, new[] { new GlyphRun(primary, glyphs, cmap) }, null);
        }

        private static void AddCodePoint(RunBuilder builder, List<ResolvedTypeface> chain, int codePoint, ref MissingCharacter? missing)
        {
            foreach (var face in chain)
            {
                var glyph = face.Typeface.GetGlyph(codePoint);
                if (glyph != 0)
                {
                    builder.Add(face, glyph, codePoint);
                    return;
                }
            }

            // どのフォントにも無い。描画を続ける設定の場合に備え、従来どおりセルのフォントの .notdef を置く。
            missing ??= new MissingCharacter(codePoint, null);
            builder.Add(chain[0], 0, null);
        }

        private static bool TryResolveVariation(
            List<ResolvedTypeface> chain, int baseCodePoint, int selector,
            out ResolvedTypeface face, out ushort glyph, out bool isDefault)
        {
            foreach (var candidate in chain)
            {
                if (!VariationSequenceTable.For(candidate.Typeface).TryLookup(baseCodePoint, selector, out isDefault, out glyph))
                {
                    continue;
                }

                if (isDefault)
                {
                    glyph = candidate.Typeface.GetGlyph(baseCodePoint);
                }

                if (glyph != 0)
                {
                    face = candidate;
                    return true;
                }
            }

            face = chain[0];
            glyph = 0;
            isDefault = false;
            return false;
        }

        private static bool VariationSelectorAt(List<int> codePoints, int index) =>
            VariationSequenceTable.IsVariationSelector(codePoints[index]);

        /// <summary>仮名 + 合成用濁点・半濁点(U+3099/U+309A)を合成済みの仮名にする(要件11.4)。</summary>
        private static string ComposeKanaVoicedMark(string cluster)
        {
            if (cluster.Length != 2 || (cluster[1] != '゙' && cluster[1] != '゚') || !IsKana(cluster[0]))
            {
                return cluster;
            }

            var composed = cluster.Normalize(NormalizationForm.FormC);
            return composed.Length == 1 ? composed : cluster;
        }

        private static bool IsKana(char c) => (c >= 'ぁ' && c <= 'ゟ') || (c >= '゠' && c <= 'ヿ');

        private static List<int> ToCodePoints(string cluster)
        {
            var list = new List<int>(2);
            foreach (var rune in cluster.EnumerateRunes())
            {
                list.Add(rune.Value);
            }

            return list;
        }

        /// <summary>同じ書体が続く字形を1つの並びにまとめる。</summary>
        private sealed class RunBuilder
        {
            private readonly List<GlyphRun> _runs = new();
            private ResolvedTypeface? _face;
            private List<ushort> _glyphs = new();
            private List<(int, ushort)> _cmap = new();

            public void Add(ResolvedTypeface face, ushort glyph, int? codePoint)
            {
                if (_face is not null && !ReferenceEquals(_face, face))
                {
                    Flush();
                }

                _face = face;
                _glyphs.Add(glyph);
                if (codePoint is { } value && glyph != 0)
                {
                    _cmap.Add((value, glyph));
                }
            }

            public IReadOnlyList<GlyphRun> Build()
            {
                Flush();
                return _runs;
            }

            private void Flush()
            {
                if (_face is null || _glyphs.Count == 0)
                {
                    return;
                }

                _runs.Add(new GlyphRun(_face, _glyphs.ToArray(), _cmap));
                _glyphs = new List<ushort>();
                _cmap = new List<(int, ushort)>();
            }
        }
    }

    /// <summary>同じ書体で続けて描く字形の並び。</summary>
    /// <param name="Face">書体(太字・斜体の合成の要否を含む)。</param>
    /// <param name="Glyphs">元の書体での字形番号。</param>
    /// <param name="CharacterMap">文字 → 字形番号の対応(サブセットの <c>cmap</c>、すなわちPDFの文字列抽出に使う)。</param>
    internal sealed record GlyphRun(ResolvedTypeface Face, ushort[] Glyphs, IReadOnlyList<(int CodePoint, ushort Glyph)> CharacterMap);

    /// <summary>どのフォントにも字形が無かった文字(要件5.5)。</summary>
    /// <param name="CodePoint">字形が無かった文字。異体字の場合は異体字セレクタ。</param>
    /// <param name="BaseCodePoint">異体字の場合の基底文字。</param>
    internal sealed record MissingCharacter(int CodePoint, int? BaseCodePoint);

    /// <summary><see cref="GlyphShaper.Shape"/> の結果。</summary>
    /// <param name="Primary">セルのフォントを解決した書体(行の高さ・下線位置の基準)。</param>
    /// <param name="Runs">字形の並び。</param>
    /// <param name="Missing">どのフォントにも字形が無かった最初の文字。全文字そろっていれば null。</param>
    internal sealed record ShapedText(ResolvedTypeface Primary, IReadOnlyList<GlyphRun> Runs, MissingCharacter? Missing);
}
