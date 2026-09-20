using SkiaSharp;

namespace Utsushi.Rendering
{
    /// <summary>
    /// 解決済みの書体と、実フォントに字形が無いため描画側で補う必要がある装飾。
    /// </summary>
    /// <param name="Typeface">実際に描画・計測に使う書体。</param>
    /// <param name="SynthesizeBold">
    /// 太字の字形を持たないため、描画側で輪郭を太らせて太字を再現する必要があるかどうか。
    /// </param>
    /// <param name="SynthesizeItalic">
    /// 斜体の字形を持たないため、描画側で傾けて斜体を再現する必要があるかどうか。
    /// </param>
    /// <remarks>
    /// <para>
    /// SkiaSharp に太字/斜体を合成させる(合成字形を持つ <see cref="SKTypeface"/> をそのまま描画する)と、
    /// PDFには <b>Type 3 フォント</b>として出力される。Type 3 になると文字列検索・コピーができず、
    /// ファイルサイズも増える。
    /// </para>
    /// <para>
    /// そこで、実フォントに字形が無い場合は「通常字形の書体をそのまま埋め込み、描画時に装飾を足す」方式を採る。
    /// こうすると埋め込みは CID TrueType のままになり、検索可能性を保ったまま太字/斜体を再現できる
    /// (実測: 合成太字の書体をそのまま使うと Type 3、通常字形+描画時の太らせでは CID TrueType)。
    /// </para>
    /// </remarks>
    public sealed record ResolvedTypeface(SKTypeface Typeface, bool SynthesizeBold, bool SynthesizeItalic);
}
