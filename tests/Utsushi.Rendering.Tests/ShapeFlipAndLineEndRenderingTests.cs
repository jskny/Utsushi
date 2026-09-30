using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// 反転した図形(要件10.17)と矢印(要件10.18)の PDF 描画の検証。
    /// </summary>
    /// <remarks>
    /// PDF のページの内容(圧縮されたコンテンツストリーム)を展開し、<c>q</c>/<c>Q</c>/<c>cm</c> を追って、塗りつぶし(<c>f</c>)と
    /// 文字(<c>BT</c>)の時点の変換行列を調べる。行列の行列式が負なら鏡像になっている。
    /// </remarks>
    public sealed class ShapeFlipAndLineEndRenderingTests : IDisposable
    {
        private static readonly FontStyle Font = new("Calibri", 10.0, false, false, UnderlineStyle.None, false, ArgbColor.Black);

        private readonly FontResolver _fontResolver;
        private readonly SkiaPdfRenderer _renderer;

        public ShapeFlipAndLineEndRenderingTests()
        {
            // 描画の構造だけを見るテストのため、フォントの代替を許容し字形の欠落は検証しない(docs/実装設計失敗事例集.md 5.2)。
            _fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            _renderer = new SkiaPdfRenderer(
                new SkiaFontMetricsProvider(_fontResolver),
                PdfRenderOptions.Default with { MissingGlyphs = MissingGlyphPolicy.Render });
        }

        public void Dispose() => _fontResolver.Dispose();

        private static ShapeCommand ArrowShape(bool flipH, bool flipV, double rotation = 0) =>
            new(
                new RectPt(100, 100, 80, 30),
                ShapePresetType.RightArrow,
                Array.Empty<double>(),
                rotation,
                new SolidShapeFill(ArgbColor.Black),
                null,
                new[] { new ShapeTextLine(new PointPt(110, 120), "AB", Font, TextAnchor.Left) },
                flipH,
                flipV);

        // -- 要件10.17: 反転 -------------------------------------------------

        [Fact]
        public void 左右反転した図形は形状だけが鏡像になり文字は変わらない()
        {
            var unflipped = Analyze(ArrowShape(false, false));
            var flipped = Analyze(ArrowShape(true, false));

            // 形状: 反転していない図形とは行列式の符号が逆になる。
            Assert.Equal(-Math.Sign(Assert.Single(unflipped.Fills).Determinant), Math.Sign(Assert.Single(flipped.Fills).Determinant));

            // 文字: 変換行列も文字の命令も反転していない図形と同じ。
            var unflippedText = Assert.Single(unflipped.Texts);
            var flippedText = Assert.Single(flipped.Texts);
            Assert.Equal(unflippedText.Matrix, flippedText.Matrix);
            Assert.Equal(unflippedText.Body, flippedText.Body);
        }

        [Theory]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void 上下反転した図形の文字は鏡像にせず180度回して描く(bool flipH, bool flipV)
        {
            // Office は上下反転した図形の文字を逆さ(180度回転)にし、鏡像にはしない(要件10.17補足)。
            var unflipped = Analyze(ArrowShape(false, false));
            var flipped = Analyze(ArrowShape(flipH, flipV));

            // 形状: 片方向の反転は行列式の符号が逆、両方向は同じ(180度回転)。
            var expectedSign = flipH ^ flipV ? -1 : 1;
            Assert.Equal(
                expectedSign * Math.Sign(Assert.Single(unflipped.Fills).Determinant),
                Math.Sign(Assert.Single(flipped.Fills).Determinant));

            // 文字: 鏡像ではない(行列式の符号が同じ)が、180度回っているため行列は異なる。
            var unflippedText = Assert.Single(unflipped.Texts);
            var flippedText = Assert.Single(flipped.Texts);
            Assert.Equal(Math.Sign(Det(unflippedText.Matrix)), Math.Sign(Det(flippedText.Matrix)));
            Assert.NotEqual(unflippedText.Matrix, flippedText.Matrix);
            Assert.Equal(unflippedText.Body, flippedText.Body);
        }

        private static double Det(Matrix m) => (m.A * m.D) - (m.B * m.C);

        [Fact]
        public void 回転と反転を併せ持つ図形でも文字は回転だけがかかる()
        {
            var rotatedOnly = Analyze(ArrowShape(false, false, rotation: 30));
            var rotatedAndFlipped = Analyze(ArrowShape(true, false, rotation: 30));

            Assert.Equal(
                -Math.Sign(Assert.Single(rotatedOnly.Fills).Determinant),
                Math.Sign(Assert.Single(rotatedAndFlipped.Fills).Determinant));
            Assert.Equal(Assert.Single(rotatedOnly.Texts).Matrix, Assert.Single(rotatedAndFlipped.Texts).Matrix);
        }

        [Fact]
        public void 反転した図形の後に描く命令は反転の影響を受けない()
        {
            // 反転のために積んだ変換を戻し忘れると、後続の塗りつぶしまで鏡像になる。
            var following = new FillRectCommand(new RectPt(300, 300, 20, 20), ArgbColor.Black);

            var baseline = Analyze(following);
            var afterFlipped = Analyze(ArrowShape(true, false), following);

            Assert.Equal(baseline.Fills.Last().Matrix, afterFlipped.Fills.Last().Matrix);
        }

        // -- 要件10.18: 矢印 -------------------------------------------------

        [Fact]
        public void 矢印付きの接続線は線に加えて矢印を塗る()
        {
            ConnectorCommand Connector(LineEndStyle? head, LineEndStyle? tail) => new(
                new RectPt(100, 100, 100, 50),
                ConnectorPresetType.Straight,
                0,
                false,
                false,
                new ShapeOutline(ArgbColor.Black, 1.0, head, tail),
                null,
                null);

            var triangle = new LineEndStyle(LineEndType.Triangle, LineEndSize.Medium, LineEndSize.Medium);
            var arrow = new LineEndStyle(LineEndType.Arrow, LineEndSize.Medium, LineEndSize.Medium);

            var none = Analyze(Connector(null, null));
            var tailOnly = Analyze(Connector(null, triangle));
            var both = Analyze(Connector(triangle, triangle));
            var open = Analyze(Connector(arrow, null));

            Assert.Empty(none.Fills);
            Assert.Single(tailOnly.Fills);
            Assert.Equal(2, both.Fills.Count);

            // arrow は開いた矢印のため塗らずに線で描く。
            Assert.Empty(open.Fills);
            Assert.Equal(none.Strokes.Count + 1, open.Strokes.Count);
        }

        [Fact]
        public void 矢印付きの線吹き出しは引き出し線の両端に矢印を塗り矢印の無い図形には描かない()
        {
            ShapeCommand Callout(ShapePresetType preset, LineEndStyle? head, LineEndStyle? tail) => new(
                new RectPt(100, 100, 80, 30),
                preset,
                Array.Empty<double>(),
                0,
                null,
                new ShapeOutline(ArgbColor.Black, 1.0, head, tail),
                Array.Empty<ShapeTextLine>());

            var oval = new LineEndStyle(LineEndType.Oval, LineEndSize.Small, LineEndSize.Small);

            Assert.Empty(Analyze(Callout(ShapePresetType.BorderCallout2, null, null)).Fills);
            Assert.Equal(2, Analyze(Callout(ShapePresetType.BorderCallout2, oval, oval)).Fills.Count);

            // 線吹き出し以外は引き出し線が無いため、矢印の指定があっても描かない。
            Assert.Empty(Analyze(Callout(ShapePresetType.Rect, oval, oval)).Fills);
        }

        [Fact]
        public void 反転した線吹き出しの矢印は形状と一緒に鏡像になる()
        {
            var triangle = new LineEndStyle(LineEndType.Triangle, LineEndSize.Medium, LineEndSize.Medium);
            var callout = new ShapeCommand(
                new RectPt(100, 100, 80, 30),
                ShapePresetType.Callout1,
                Array.Empty<double>(),
                0,
                null,
                new ShapeOutline(ArgbColor.Black, 1.0, null, triangle),
                Array.Empty<ShapeTextLine>());

            var unflipped = Assert.Single(Analyze(callout).Fills);
            var flipped = Assert.Single(Analyze(callout with { FlipHorizontal = true }).Fills);

            Assert.Equal(-Math.Sign(unflipped.Determinant), Math.Sign(flipped.Determinant));
        }

        // -- PDF の解析 -------------------------------------------------

        private PageContent Analyze(params DrawCommand[] commands)
        {
            var page = new PageLayout(
                PaperSize.A4,
                PageOrientation.Portrait,
                PaperSize.A4.WidthPt,
                PaperSize.A4.HeightPt,
                commands,
                1,
                (1, 1),
                (1, 1),
                1.0);

            using var output = new MemoryStream();
            _renderer.Render(new PagedLayout(new[] { page }, "test-report", "テストシート"), output);
            return PageContent.Parse(ReadContentStreams(output.ToArray()));
        }

        /// <summary>PDF 中の Flate 圧縮されたストリームのうち、描画命令を含むもの(ページの内容)を展開して連結する。</summary>
        private static string ReadContentStreams(byte[] pdf)
        {
            var text = Encoding.Latin1.GetString(pdf);
            var result = new StringBuilder();
            foreach (Match match in Regex.Matches(text, @"stream\r?\n"))
            {
                var start = match.Index + match.Length;
                var end = text.IndexOf("endstream", start, StringComparison.Ordinal);
                if (end < 0 || end - start < 2)
                {
                    continue;
                }

                string decoded;
                try
                {
                    // zlib のヘッダー(2バイト)を読み飛ばして Deflate として展開する。
                    using var deflate = new DeflateStream(new MemoryStream(pdf, start + 2, end - start - 2), CompressionMode.Decompress);
                    using var buffer = new MemoryStream();
                    deflate.CopyTo(buffer);
                    decoded = Encoding.Latin1.GetString(buffer.ToArray());
                }
                catch (InvalidDataException)
                {
                    continue; // 圧縮されていないストリーム・フォント等
                }

                if (Regex.IsMatch(decoded, @"(^|\s)cm(\s|$)"))
                {
                    result.Append(decoded).Append('\n');
                }
            }

            return result.ToString();
        }

        /// <summary>変換行列 [a b c d e f](PDF の <c>cm</c> と同じ並び)。比較は小数第3位で丸める。</summary>
        private readonly struct Matrix : IEquatable<Matrix>
        {
            public Matrix(double a, double b, double c, double d, double e, double f)
            {
                A = a;
                B = b;
                C = c;
                D = d;
                E = e;
                F = f;
            }

            public static Matrix Identity => new(1, 0, 0, 1, 0, 0);

            public double A { get; }

            public double B { get; }

            public double C { get; }

            public double D { get; }

            public double E { get; }

            public double F { get; }

            /// <summary>PDF の <c>cm</c>: CTM' = M × CTM。</summary>
            public Matrix Concat(Matrix m) => new(
                (m.A * A) + (m.B * C),
                (m.A * B) + (m.B * D),
                (m.C * A) + (m.D * C),
                (m.C * B) + (m.D * D),
                (m.E * A) + (m.F * C) + E,
                (m.E * B) + (m.F * D) + F);

            public bool Equals(Matrix other) => ToString() == other.ToString();

            public override bool Equals(object? obj) => obj is Matrix other && Equals(other);

            public override int GetHashCode() => ToString().GetHashCode(StringComparison.Ordinal);

            public override string ToString() => string.Join(
                " ", new[] { A, B, C, D, E, F }.Select(v => Math.Round(v, 3).ToString("0.###", CultureInfo.InvariantCulture)));
        }

        private sealed record PaintOperation(Matrix Matrix)
        {
            public double Determinant => (Matrix.A * Matrix.D) - (Matrix.B * Matrix.C);
        }

        private sealed record TextObject(Matrix Matrix, string Body);

        private sealed class PageContent
        {
            public List<PaintOperation> Fills { get; } = new();

            public List<PaintOperation> Strokes { get; } = new();

            public List<TextObject> Texts { get; } = new();

            public static PageContent Parse(string content)
            {
                var result = new PageContent();
                var stack = new Stack<Matrix>();
                var ctm = Matrix.Identity;
                var operands = new List<double>();
                StringBuilder? textBody = null;
                var textMatrix = Matrix.Identity;

                foreach (var token in content.Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (textBody is not null && token != "ET")
                    {
                        textBody.Append(token).Append(' ');
                    }

                    if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    {
                        operands.Add(number);
                        continue;
                    }

                    switch (token)
                    {
                        case "q":
                            stack.Push(ctm);
                            break;
                        case "Q":
                            ctm = stack.Pop();
                            break;
                        case "cm" when operands.Count >= 6:
                            var o = operands.Skip(operands.Count - 6).ToArray();
                            ctm = ctm.Concat(new Matrix(o[0], o[1], o[2], o[3], o[4], o[5]));
                            break;
                        case "f":
                        case "f*":
                        case "B":
                        case "B*":
                            result.Fills.Add(new PaintOperation(ctm));
                            break;
                        case "S":
                            result.Strokes.Add(new PaintOperation(ctm));
                            break;
                        case "BT":
                            textBody = new StringBuilder();
                            textMatrix = ctm;
                            break;
                        case "ET":
                            result.Texts.Add(new TextObject(textMatrix, textBody?.ToString() ?? string.Empty));
                            textBody = null;
                            break;
                    }

                    operands.Clear();
                }

                return result;
            }
        }
    }
}
