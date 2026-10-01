using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Utsushi.Rendering;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// 実フォントの計測(<see cref="SkiaFontMetricsProvider"/>)でも、折り返し(<see cref="TextWrapper"/>)が
    /// 1文字足すごとに行全体を測り直す以前の方法と同じ位置で改行することを確かめる。
    /// </summary>
    /// <remarks>
    /// 実フォントでは、字形の並び(外字用の代替フォントへの切り替え・異体字・合成用濁点)ごとに単精度で送り幅を合計するため、
    /// 文字ごとの幅の合計と行全体の幅は末尾の桁まで一致するとは限らない。それでも改行位置が変わらないことを確かめる。
    /// </remarks>
    public sealed class SkiaTextWrapEquivalenceTests : IDisposable
    {
        private static readonly string[] Pieces =
        {
            "あ", "い", "漢", "字", "株", "式", "会", "社", "東", "京", "a", "b", "W", "i", " ", "1", "0", ",", ".", "、", "。",
            "𠮷", "葛\U000E0100", "が", "é", "\t", "ｱ", "ｶﾞ", "（", "）", "�", "Ｗ", "-", "—", "👍",
        };

        private readonly FontResolver _fontResolver = new(FontResolverOptions.AllowFallback());

        public void Dispose() => _fontResolver.Dispose();

        private static List<string> ReferenceWrap(IFontMetricsProvider metrics, FontStyle font, string text, double availableWidthPt)
        {
            var lines = new List<string>();
            var unified = text.IndexOf('\r') < 0 ? text : text.Replace("\r\n", "\n").Replace('\r', '\n');
            foreach (var paragraph in unified.Split('\n'))
            {
                if (paragraph.Length == 0)
                {
                    lines.Add(string.Empty);
                    continue;
                }

                var current = new StringBuilder();
                var elements = StringInfo.GetTextElementEnumerator(paragraph);
                while (elements.MoveNext())
                {
                    var element = elements.GetTextElement();
                    if (current.Length > 0
                        && metrics.MeasureTextWidth(font, current.ToString() + element) > availableWidthPt)
                    {
                        lines.Add(current.ToString());
                        current.Clear();
                    }

                    current.Append(element);
                }

                lines.Add(current.ToString());
            }

            return lines.Count == 0 ? new List<string> { string.Empty } : lines;
        }

        [Fact]
        public void 実フォントの計測でも以前と同じ位置で改行する()
        {
            var metrics = new SkiaFontMetricsProvider(_fontResolver);
            var random = new Random(20261001);
            for (var i = 0; i < 1500; i++)
            {
                var font = FontStyle.Default with
                {
                    Name = i % 2 == 0 ? "MS PGothic" : "Calibri",
                    SizePt = 6 + (random.Next(0, 30) * 0.37),
                    Bold = i % 5 == 0,
                    Italic = i % 7 == 0,
                };

                var builder = new StringBuilder();
                var length = random.Next(0, 160);
                for (var j = 0; j < length; j++)
                {
                    builder.Append(random.Next(80) == 0 ? "\n" : Pieces[random.Next(Pieces.Length)]);
                }

                var text = builder.ToString();
                var available = random.Next(6) == 0 ? random.NextDouble() * 5 : random.NextDouble() * 300;

                var expected = ReferenceWrap(metrics, font, text, available);
                var actual = TextWrapper.Wrap(metrics, font, text, available);

                Assert.True(
                    string.Join("\u0001", expected) == string.Join("\u0001", actual),
                    $"改行位置が異なる: font={font.Name} {font.SizePt}pt, available={available}, text={text}");
            }
        }
    }
}
