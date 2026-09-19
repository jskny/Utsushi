using System;
using Utsushi.Parsing.Model;

namespace Utsushi.Layout.Text;

/// <summary>
/// 実フォントを参照せず、文字種ごとの固定比率で幅を見積もる <see cref="IFontMetricsProvider"/>。
/// </summary>
/// <remarks>
/// <para>
/// 実運用の描画には Rendering レイヤーの SkiaSharp 実装を使うこと。本実装は
/// 「実行環境のフォント構成に左右されない決定的な結果」が必要な場面
/// (レイアウトのユニットテスト・ゴールデンテスト)向けである。
/// </para>
/// <para>
/// 全角(CJK)を1em、半角を0.5emとして扱う。実フォントとは一致しないため、
/// 帳票の見た目の検証には使わないこと。
/// </para>
/// </remarks>
public sealed class ApproximateFontMetricsProvider : IFontMetricsProvider
{
    /// <summary>半角文字の幅(em単位)。</summary>
    private const double HalfWidthEm = 0.5;

    /// <summary>全角文字の幅(em単位)。</summary>
    private const double FullWidthEm = 1.0;

    private const double AscentRatio = 0.88;
    private const double DescentRatio = 0.22;
    private const double LineSpacingRatio = 1.2;

    /// <inheritdoc />
    public FontMetrics GetMetrics(FontStyle font)
    {
        var size = font.SizePt;
        return new FontMetrics(size * AscentRatio, size * DescentRatio, size * LineSpacingRatio);
    }

    /// <inheritdoc />
    public double MeasureTextWidth(FontStyle font, string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return 0.0;
        }

        var em = 0.0;
        foreach (var c in text)
        {
            em += IsFullWidth(c) ? FullWidthEm : HalfWidthEm;
        }

        var width = em * font.SizePt;

        // 太字はわずかに送り幅が広がる。
        return font.Bold ? width * 1.03 : width;
    }

    /// <summary>
    /// 全角(East Asian Wide / Fullwidth)として扱う文字かどうかを判定する。
    /// </summary>
    private static bool IsFullWidth(char c) =>
        c >= 0x1100 &&
        (c <= 0x115F                        // ハングル字母
         || (c >= 0x2E80 && c <= 0xA4CF)    // CJK部首補助〜漢字・かな・記号
         || (c >= 0xAC00 && c <= 0xD7A3)    // ハングル音節
         || (c >= 0xF900 && c <= 0xFAFF)    // CJK互換漢字
         || (c >= 0xFE30 && c <= 0xFE6F)    // CJK互換形・小字形
         || (c >= 0xFF00 && c <= 0xFF60)    // 全角英数・記号
         || (c >= 0xFFE0 && c <= 0xFFE6));  // 全角通貨記号等
}
