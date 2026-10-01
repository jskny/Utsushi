using System;
using System.Collections.Generic;
using System.Globalization;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Xunit;

namespace Utsushi.Layout.Tests
{
    /// <summary>
    /// 罫線の重複判定キー(<see cref="BorderLineKey"/>)が、以前の文字列のキー(書式 <c>"0.###"</c>)と
    /// 同じ値どうしを同じキーとみなすことを確かめる。
    /// </summary>
    public sealed class BorderLineKeyTests
    {
        /// <summary>以前のキー。</summary>
        private static string ReferenceKey(PointPt from, PointPt to, ArgbColor color, double widthPt, LineDashStyle dash) =>
            string.Format(
                CultureInfo.InvariantCulture,
                "{0:0.###},{1:0.###},{2:0.###},{3:0.###},{4},{5:0.###},{6}",
                from.X, from.Y, to.X, to.Y, color, widthPt, dash);

        private static IEnumerable<double> SampleValues()
        {
            var random = new Random(20261001);
            for (var i = 0; i < 20000; i++)
            {
                yield return (i % 5) switch
                {
                    0 => random.NextDouble() * 1200,
                    1 => Math.Round(random.NextDouble() * 1200, 3) + 0.0005, // 丸めの境目
                    2 => (random.Next(0, 1200000) + 0.5) / 1000.0, // 丸めの境目(2進数で表せない値)
                    3 => Math.Round(random.NextDouble() * 1200, 3) + ((random.NextDouble() - 0.5) * 1e-9), // 境目のごく近く
                    _ => (random.NextDouble() - 0.5) * 0.004, // 0の近く(負の値を含む)
                };
            }

            yield return 0.0;
            yield return -0.0;
            yield return 1.0005;
            yield return -1.0005;
            yield return 0.75;
        }

        [Fact]
        public void 丸めた値の同一性が以前の文字列のキーと一致する()
        {
            // 以前のキーの文字列 → 新しいキー、新しいキー → 以前のキーの文字列が、それぞれ1対1に対応すること。
            var byReference = new Dictionary<string, BorderLineKey>();
            var byKey = new Dictionary<BorderLineKey, string>();
            foreach (var value in SampleValues())
            {
                foreach (var point in new[] { new PointPt(value, 10.0), new PointPt(10.0, value) })
                {
                    var to = new PointPt(point.X + 50.0, point.Y);
                    var reference = ReferenceKey(point, to, ArgbColor.Black, value, LineDashStyle.Solid);
                    var key = new BorderLineKey(point, to, ArgbColor.Black, value, LineDashStyle.Solid);

                    if (byReference.TryGetValue(reference, out var existingKey))
                    {
                        Assert.True(existingKey.Equals(key), $"同じ文字列のキーなのに別のキーになった: {value:R}");
                    }
                    else
                    {
                        byReference[reference] = key;
                    }

                    if (byKey.TryGetValue(key, out var existingReference))
                    {
                        Assert.True(existingReference == reference, $"別の文字列のキーなのに同じキーになった: {value:R}");
                    }
                    else
                    {
                        byKey[key] = reference;
                    }
                }
            }
        }

        [Fact]
        public void 色と線種が違えば別のキーになる()
        {
            var from = new PointPt(1, 2);
            var to = new PointPt(3, 2);
            var key = new BorderLineKey(from, to, ArgbColor.Black, 0.75, LineDashStyle.Solid);

            Assert.Equal(key, new BorderLineKey(from, to, ArgbColor.Black, 0.75, LineDashStyle.Solid));
            Assert.NotEqual(key, new BorderLineKey(from, to, new ArgbColor(255, 255, 0, 0), 0.75, LineDashStyle.Solid));
            Assert.NotEqual(key, new BorderLineKey(from, to, ArgbColor.Black, 0.75, LineDashStyle.Dash));
            Assert.NotEqual(key, new BorderLineKey(from, to, ArgbColor.Black, 0.7505, LineDashStyle.Solid));
            Assert.Equal(key, new BorderLineKey(from, to, ArgbColor.Black, 0.7504, LineDashStyle.Solid));
        }
    }
}
