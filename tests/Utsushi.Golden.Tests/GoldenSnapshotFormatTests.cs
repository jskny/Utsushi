using System;
using Utsushi.Core;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Golden.Tests
{
    /// <summary>
    /// ゴールデンスナップショットのフォーマッタが、反転(要件10.17)・矢印(要件10.18)・調整値(要件10.13)を出力することの検証。
    /// </summary>
    /// <remarks>
    /// docs/実装設計失敗事例集.md 5.1(フォーマッタが回転角を出力せず、回転のバグをゴールデンテストが検出できなかった)の再発防止。
    /// 既存の期待値を変えないよう、指定が無い場合は何も出力しないことも確かめる。
    /// </remarks>
    public sealed class GoldenSnapshotFormatTests
    {
        private static string Snapshot(params DrawCommand[] commands) =>
            GoldenSnapshot.Create(new PagedLayout(
                new[]
                {
                    new PageLayout(
                        PaperSize.A4, PageOrientation.Portrait, PaperSize.A4.WidthPt, PaperSize.A4.HeightPt,
                        commands, 1, (1, 1), (1, 1), 1.0),
                },
                "test-report",
                "テストシート"));

        private static ShapeCommand Shape(
            bool flipH = false, bool flipV = false, ShapeOutline? outline = null, double[]? adjustments = null) =>
            new(
                new RectPt(10, 10, 40, 20),
                ShapePresetType.Rect,
                adjustments ?? Array.Empty<double>(),
                0,
                null,
                outline,
                Array.Empty<ShapeTextLine>(),
                flipH,
                flipV);

        [Fact]
        public void 反転していない図形の行は反転を出力しない()
        {
            Assert.DoesNotContain("flipH", Snapshot(Shape()));
        }

        [Theory]
        [InlineData(true, false, "flipH=True flipV=False")]
        [InlineData(false, true, "flipH=False flipV=True")]
        [InlineData(true, true, "flipH=True flipV=True")]
        public void 反転した図形の行は反転を出力する(bool flipH, bool flipV, string expected)
        {
            Assert.Contains(expected, Snapshot(Shape(flipH, flipV)));
        }

        [Fact]
        public void 反転の向きが違う図形は異なる行になる()
        {
            Assert.NotEqual(Snapshot(Shape(true, false)), Snapshot(Shape(false, true)));
            Assert.NotEqual(Snapshot(Shape()), Snapshot(Shape(true, false)));
        }

        [Fact]
        public void 矢印の無い枠線は従来どおりの書式になる()
        {
            var snapshot = Snapshot(Shape(outline: new ShapeOutline(ArgbColor.Black, 1.0)));

            Assert.Contains("outline=#FF000000/1pt ", snapshot);
            Assert.DoesNotContain("head=", snapshot);
            Assert.DoesNotContain("tail=", snapshot);
        }

        [Fact]
        public void 図形と接続線の枠線の矢印を種類と大きさつきで出力する()
        {
            var outline = new ShapeOutline(
                ArgbColor.Black,
                1.0,
                new LineEndStyle(LineEndType.Oval, LineEndSize.Small, LineEndSize.Large),
                new LineEndStyle(LineEndType.Triangle, LineEndSize.Medium, LineEndSize.Medium));
            var connector = new ConnectorCommand(
                new RectPt(10, 50, 40, 20), ConnectorPresetType.Straight, 0, false, false, outline, null, null);

            var lines = Snapshot(Shape(outline: outline), connector).Split('\n');

            const string expected = "outline=#FF000000/1pt/head=Oval:Small:Large/tail=Triangle:Medium:Medium ";
            Assert.Contains(lines, l => l.StartsWith("shape", StringComparison.Ordinal) && l.Contains(expected, StringComparison.Ordinal));
            Assert.Contains(lines, l => l.StartsWith("connector", StringComparison.Ordinal) && l.Contains(expected, StringComparison.Ordinal));
        }

        [Fact]
        public void 矢印の種類が違う枠線は異なる行になる()
        {
            ShapeOutline Outline(LineEndType type) =>
                new(ArgbColor.Black, 1.0, null, new LineEndStyle(type, LineEndSize.Medium, LineEndSize.Medium));

            Assert.NotEqual(
                Snapshot(Shape(outline: Outline(LineEndType.Triangle))),
                Snapshot(Shape(outline: Outline(LineEndType.Stealth))));
        }

        [Fact]
        public void 調整値が無いか全て未指定の図形の行は調整値を出力しない()
        {
            Assert.DoesNotContain("adj=", Snapshot(Shape()));
            Assert.DoesNotContain("adj=", Snapshot(Shape(adjustments: new[] { double.NaN, double.NaN, double.NaN, double.NaN })));
        }

        [Fact]
        public void 調整値の指定がある図形の行は未指定を含めて調整値を出力する()
        {
            var snapshot = Snapshot(Shape(adjustments: new[] { 0.5, 1.0, double.NaN, 1.6 }));

            Assert.Contains(" adj=[0.5,1,-,1.6]", snapshot);
        }

        [Fact]
        public void 調整値が違う図形は異なる行になる()
        {
            Assert.NotEqual(
                Snapshot(Shape(adjustments: new[] { 0.1875, -0.08333, 1.125, -0.38333 })),
                Snapshot(Shape(adjustments: new[] { 0.5, 1.0, 1.5, 1.6 })));
        }
    }
}
