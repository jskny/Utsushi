using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Utsushi.Layout.Text;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Layout.Tests
{
    /// <summary>
    /// 折り返し(<see cref="TextWrapper"/>)が、1文字足すごとに行全体を測り直す素直な方法と
    /// 同じ位置で改行すること、および計測回数が行の長さの2乗に比例しないことを確かめる。
    /// </summary>
    public sealed class TextWrapperTests
    {
        private static readonly string[] Pieces =
        {
            "あ", "い", "漢", "字", "株", "式", "会", "社", "a", "b", "W", "i", " ", "1", ",", ".", "、", "。",
            "𠮷", "葛\U000E0100", "が", "é", "\t", "ｱ", "ｶﾞ", "（", "）", "�", "Ｗ", "-", "👍",
        };

        /// <summary>以前の実装(1文字足すごとに行全体を測り直す)。比較の基準にする。</summary>
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

        private static string RandomText(Random random, int length)
        {
            var builder = new StringBuilder();
            for (var i = 0; i < length; i++)
            {
                if (random.Next(60) == 0)
                {
                    builder.Append(random.Next(2) == 0 ? "\n" : "\r\n");
                    continue;
                }

                builder.Append(Pieces[random.Next(Pieces.Length)]);
            }

            return builder.ToString();
        }

        public static IEnumerable<object[]> Providers()
        {
            yield return new object[] { "approximate" };
            yield return new object[] { "irregular" };
        }

        private static IFontMetricsProvider CreateProvider(string name) =>
            name == "approximate" ? new ApproximateFontMetricsProvider() : new IrregularFontMetricsProvider();

        [Theory]
        [MemberData(nameof(Providers))]
        public void 一文字ずつ測り直す方法と同じ位置で改行する(string providerName)
        {
            var provider = CreateProvider(providerName);
            var random = new Random(20261001);
            for (var i = 0; i < 3000; i++)
            {
                var font = FontStyle.Default with { SizePt = 6 + (random.Next(0, 30) * 0.37), Bold = i % 5 == 0 };
                var text = RandomText(random, random.Next(0, 200));
                var available = random.Next(6) switch
                {
                    0 => (random.NextDouble() * 5) - 1, // 1文字も収まらない幅・負の幅
                    1 => 1e9, // 折り返さない
                    _ => random.NextDouble() * 300,
                };

                var expected = ReferenceWrap(provider, font, text, available);
                var actual = TextWrapper.Wrap(provider, font, text, available);

                Assert.True(
                    string.Join("\u0001", expected) == string.Join("\u0001", actual),
                    $"改行位置が異なる: available={available}, text={text}");
            }
        }

        [Theory]
        [InlineData(double.NaN)]
        [InlineData(double.PositiveInfinity)]
        [InlineData(double.NegativeInfinity)]
        [InlineData(0.0)]
        public void 特殊な幅でも以前と同じ結果になる(double available)
        {
            var provider = new ApproximateFontMetricsProvider();
            const string text = "株式会社サンプル 御中\nabc";

            Assert.Equal(
                ReferenceWrap(provider, FontStyle.Default, text, available),
                TextWrapper.Wrap(provider, FontStyle.Default, text, available));
        }

        [Fact]
        public void 一行に収まる文字列は一回の計測で済む()
        {
            var provider = new CountingFontMetricsProvider(new ApproximateFontMetricsProvider());
            var text = new string('あ', 2000);

            var lines = TextWrapper.Wrap(provider, FontStyle.Default, text, 1e9);

            Assert.Single(lines);
            Assert.Equal(1, provider.Calls);
        }

        [Theory]
        [MemberData(nameof(Providers))]
        public void 長い段落を折り返しても計測回数は行数に比例する程度に収まる(string providerName)
        {
            var provider = new CountingFontMetricsProvider(CreateProvider(providerName));
            var random = new Random(7);
            var builder = new StringBuilder();
            for (var i = 0; i < 2000; i++)
            {
                builder.Append(Pieces[random.Next(Pieces.Length)].Replace("\t", "a"));
            }

            var lines = TextWrapper.Wrap(provider, FontStyle.Default, builder.ToString(), 120.0);

            // 以前の実装は1文字ごとに1回(約2000回)計測していた。
            Assert.True(lines.Count > 50, $"行数: {lines.Count}");
            Assert.True(provider.Calls <= (lines.Count * 6) + 10, $"計測回数: {provider.Calls}, 行数: {lines.Count}");
        }

        /// <summary>
        /// 文字ごとに不規則な幅(0を含む)を持ち、同じ文字が続くと詰める(カーニングに相当)ために
        /// 文字ごとの幅の合計と文字列全体の幅が一致しない計測。文字を足しても幅は減らない。
        /// </summary>
        private sealed class IrregularFontMetricsProvider : IFontMetricsProvider
        {
            public FontMetrics GetMetrics(FontStyle font) => new(font.SizePt * 0.8, font.SizePt * 0.2, font.SizePt * 1.2);

            public double MeasureTextWidth(FontStyle font, string text)
            {
                var width = 0.0;
                var previous = -1;
                foreach (var rune in text.EnumerateRunes())
                {
                    var advance = (((uint)rune.Value * 2654435761u) % 97u) / 10.0 * (font.SizePt / 10.0);
                    if (rune.Value == previous)
                    {
                        advance -= Math.Min(0.3, advance);
                    }

                    width += advance;
                    previous = rune.Value;
                }

                // 単精度に丸める(実フォントの計測と同様、合計の順序で末尾の桁が変わる)。
                return (float)width;
            }
        }

        private sealed class CountingFontMetricsProvider : IFontMetricsProvider
        {
            private readonly IFontMetricsProvider _inner;

            public CountingFontMetricsProvider(IFontMetricsProvider inner) => _inner = inner;

            public int Calls { get; private set; }

            public FontMetrics GetMetrics(FontStyle font) => _inner.GetMetrics(font);

            public double MeasureTextWidth(FontStyle font, string text)
            {
                Calls++;
                return _inner.MeasureTextWidth(font, text);
            }
        }
    }
}
