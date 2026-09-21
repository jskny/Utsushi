using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Layout.Model;
using Utsushi.Parsing.Model;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// PDF描画の検証(要件5.1-5.4、タスク6.1-6.3)。
    /// </summary>
    public sealed class SkiaPdfRendererTests : IDisposable
    {
        private readonly FontResolver _fontResolver;
        private readonly SkiaFontMetricsProvider _metrics;

        public SkiaPdfRendererTests()
        {
            // CI/開発環境に特定フォントがあるとは限らないため、描画テストではフォールバックを許容する。
            _fontResolver = new FontResolver(FontResolverOptions.AllowFallback());
            _metrics = new SkiaFontMetricsProvider(_fontResolver);
        }

        public void Dispose() => _fontResolver.Dispose();

        private static PagedLayout Layout(int pageCount = 1, IReadOnlyList<DrawCommand>? commands = null)
        {
            var pages = new List<PageLayout>();
            for (var i = 1; i <= pageCount; i++)
            {
                pages.Add(new PageLayout(
                    PaperSize.A4,
                    PageOrientation.Portrait,
                    PaperSize.A4.WidthPt,
                    PaperSize.A4.HeightPt,
                    commands ?? DefaultCommands(),
                    i,
                    (1, 10),
                    (1, 5),
                    1.0));
            }

            return new PagedLayout(pages, "test-report", "テストシート");
        }

        private static IReadOnlyList<DrawCommand> DefaultCommands() => new DrawCommand[]
        {
        new FillRectCommand(new RectPt(50, 50, 200, 30), new ArgbColor(0xFF, 0xDC, 0xE6, 0xF1)),
        new LineCommand(new PointPt(50, 50), new PointPt(250, 50), ArgbColor.Black, 0.75, LineDashStyle.Solid),
        new TextCommand(new PointPt(55, 70), "テスト", FontStyle.Default, TextAnchor.Left, null),
        };

        [Fact]
        public void 単一のPDFファイルとして出力される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            renderer.Render(Layout(pageCount: 3), output);

            var bytes = output.ToArray();
            Assert.True(bytes.Length > 0);
            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));

            // 3ページぶんの Page オブジェクトが含まれる
            var content = Encoding.Latin1.GetString(bytes);
            Assert.Contains("/Type /Page", content);
        }

        [Fact]
        public void ページサイズは用紙と向きに対応する()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var landscape = new PagedLayout(
                new[]
                {
                new PageLayout(
                    PaperSize.A4, PageOrientation.Landscape,
                    PaperSize.A4.HeightPt, PaperSize.A4.WidthPt,
                    DefaultCommands(), 1, (1, 1), (1, 1), 1.0),
                },
                "test-report", "テストシート");

            renderer.Render(landscape, output);

            var content = Encoding.Latin1.GetString(output.ToArray());
            var mediaBox = System.Text.RegularExpressions.Regex.Match(
                content, @"/MediaBox\s*\[\s*([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s*\]");

            Assert.True(mediaBox.Success, "PDF に /MediaBox が含まれるはず");

            var width = double.Parse(mediaBox.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture);
            var height = double.Parse(mediaBox.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);

            // A4横 = 841.89 x 595.28 pt。PDF側の数値は有効桁が丸められるため許容誤差を持たせる。
            Assert.Equal(PaperSize.A4.HeightPt, width, precision: 0);
            Assert.Equal(PaperSize.A4.WidthPt, height, precision: 0);
            Assert.True(width > height, "横向きなので幅のほうが大きいはず");
        }

        [Fact]
        public void ページが無ければエラーになる()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var ex = Assert.Throws<PdfRenderingException>(
                () => renderer.Render(new PagedLayout(Array.Empty<PageLayout>(), "r", "s"), output));

            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal(0, output.Length);
        }

        [Fact]
        public void アウトライン出力はフォントを埋め込まない()
        {
            using var embedded = new MemoryStream();
            using var outlined = new MemoryStream();

            new SkiaPdfRenderer(_metrics, PdfRenderOptions.Default).Render(Layout(), embedded);
            new SkiaPdfRenderer(_metrics, PdfRenderOptions.OutlineText).Render(Layout(), outlined);

            // フォントを埋め込まないぶん、アウトライン出力のほうが小さくなる。
            Assert.True(
                outlined.Length < embedded.Length,
                $"アウトライン出力のほうが小さいはず (outline={outlined.Length}, embed={embedded.Length})");

            var content = Encoding.Latin1.GetString(outlined.ToArray());
            Assert.DoesNotContain("/FontFile", content);
        }

        [Fact]
        public void ファイル出力は一時ファイル経由で確定する()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            var path = Path.Combine(directory, "out.pdf");

            try
            {
                renderer.RenderToFile(Layout(), path);

                Assert.True(File.Exists(path));
                Assert.False(File.Exists(path + ".utsushi-tmp"), "一時ファイルが残ってはいけない");
                Assert.True(new FileInfo(path).Length > 0);
            }
            finally
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
        }

        [Fact]
        public void 描画に失敗しても出力先には何も書き込まない()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            // 未知の描画命令を1つ混ぜて途中で失敗させる。
            var layout = Layout(commands: new DrawCommand[]
            {
            new FillRectCommand(new RectPt(0, 0, 10, 10), ArgbColor.Black),
            new UnknownCommand(),
            });

            Assert.Throws<PdfRenderingException>(() => renderer.Render(layout, output));

            // 要件5.4: 不完全なPDFを出力先に残さない
            Assert.Equal(0, output.Length);
        }

        [Fact]
        public void 既存ファイルがあっても失敗時は上書きされない()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, "existing.pdf");

            try
            {
                File.WriteAllText(path, "既存の内容");

                var layout = Layout(commands: new DrawCommand[] { new UnknownCommand() });
                Assert.Throws<PdfRenderingException>(() => renderer.RenderToFile(layout, path));

                Assert.Equal("既存の内容", File.ReadAllText(path));
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void 太字を含むPDFでもType3フォントにならない()
        {
            // 太字の字形を持たないフォントで太字を描くと、SkiaSharp に合成させた場合は
            // PDF が Type 3 フォントになり文字列検索ができなくなる。
            // Utsushi は通常字形を埋め込んで描画時に輪郭を太らせるため、Type 3 にならない。
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var bold = FontStyle.Default with { Bold = true };
            var layout = Layout(commands: new DrawCommand[]
            {
            new TextCommand(new PointPt(55, 70), "太字テスト Bold", bold, TextAnchor.Left, null),
            new TextCommand(new PointPt(55, 90), "通常テスト", FontStyle.Default, TextAnchor.Left, null),
            });

            renderer.Render(layout, output);

            var content = Encoding.Latin1.GetString(output.ToArray());
            Assert.DoesNotContain("/Subtype /Type3", content);
            Assert.Contains("/FontFile2", content);
        }

        [Fact]
        public void 斜体を含むPDFでもType3フォントにならない()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var italic = FontStyle.Default with { Italic = true };
            var layout = Layout(commands: new DrawCommand[]
            {
            new TextCommand(new PointPt(55, 70), "斜体テスト Italic", italic, TextAnchor.Left, null),
            });

            renderer.Render(layout, output);

            var content = Encoding.Latin1.GetString(output.ToArray());
            Assert.DoesNotContain("/Subtype /Type3", content);
        }

        [Fact]
        public void 画像はPDFに描画される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ImageCommand(new RectPt(10, 10, 60, 20), TinyPng(), "image/png"),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
            var content = Encoding.Latin1.GetString(output.ToArray());
            Assert.Contains("/Image", content);
        }

        [Fact]
        public void デコードできない画像は帳票コードとシート名を含む例外になる()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ImageCommand(new RectPt(0, 0, 10, 10), new byte[] { 0x00, 0x01, 0x02 }, "image/png"),
            });

            var ex = Assert.Throws<PdfRenderingException>(() => renderer.Render(layout, output));

            Assert.Equal(ProcessingStage.Rendering, ex.Stage);
            Assert.Equal("test-report", ex.ReportCode);
            Assert.Equal("テストシート", ex.SheetName);
            Assert.Equal(0, output.Length);
        }

        [Fact]
        public void 宣言サイズが上限を超える画像はデコード前に拒否される()
        {
            // いわゆるピクセル爆弾対策(security-reviewer指摘)。実際のピクセルデータが
            // 無い/不正な状態でも、IHDRの宣言サイズだけで数百バイトのファイルが
            // 数億ピクセル相当を要求できてしまうため、SKBitmap.Decodeで実際に展開する前に
            // SKBitmap.DecodeBoundsで寸法を確認し拒否する。
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var hugePng = BuildPngWithDeclaredSize(20000, 20000);
            var layout = Layout(commands: new DrawCommand[]
            {
                new ImageCommand(new RectPt(0, 0, 10, 10), hugePng, "image/png"),
            });

            Assert.True(hugePng.Length < 200, "この検証はファイルサイズが小さいことが前提(実データを展開させないため)");
            Assert.Throws<PdfRenderingException>(() => renderer.Render(layout, output));
            Assert.Equal(0, output.Length);
        }

        [Fact]
        public void 図形はPDFに描画される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ShapeCommand(
                    new RectPt(10, 10, 60, 20),
                    ShapePresetType.Rect,
                    Array.Empty<double>(),
                    RotationDegrees: 0,
                    Fill: new SolidShapeFill(ArgbColor.Black),
                    Outline: new ShapeOutline(ArgbColor.Black, 1.0),
                    TextLines: Array.Empty<ShapeTextLine>()),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        [Theory]
        [InlineData(ShapePresetType.Rect)]
        [InlineData(ShapePresetType.RoundRect)]
        [InlineData(ShapePresetType.Ellipse)]
        [InlineData(ShapePresetType.Triangle)]
        [InlineData(ShapePresetType.RightArrow)]
        [InlineData(ShapePresetType.LeftArrow)]
        [InlineData(ShapePresetType.UpArrow)]
        [InlineData(ShapePresetType.DownArrow)]
        [InlineData(ShapePresetType.LeftRightArrow)]
        [InlineData(ShapePresetType.UpDownArrow)]
        [InlineData(ShapePresetType.WedgeRectCallout)]
        [InlineData(ShapePresetType.WedgeRoundRectCallout)]
        [InlineData(ShapePresetType.WedgeEllipseCallout)]
        public void 全プリセットが例外なく描画できる(ShapePresetType preset)
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ShapeCommand(
                    new RectPt(10, 10, 60, 40),
                    preset,
                    Array.Empty<double>(),
                    RotationDegrees: 0,
                    Fill: new SolidShapeFill(ArgbColor.Black),
                    Outline: null,
                    TextLines: Array.Empty<ShapeTextLine>()),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void 回転した図形の後にも他の描画命令が正しく描画される()
        {
            // 図形の回転はcanvas.Save/RotateDegrees/Restoreで実装しており、
            // Restore漏れがあると後続の描画命令の座標系がずれてしまう回帰テスト。
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ShapeCommand(
                    new RectPt(10, 10, 40, 40),
                    ShapePresetType.Rect,
                    Array.Empty<double>(),
                    RotationDegrees: 30,
                    Fill: new SolidShapeFill(ArgbColor.Black),
                    Outline: null,
                    TextLines: Array.Empty<ShapeTextLine>()),
                new FillRectCommand(new RectPt(100, 100, 30, 30), ArgbColor.Black),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void グラデーション塗りの図形はPDFに描画される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ShapeCommand(
                    new RectPt(10, 10, 60, 20),
                    ShapePresetType.Rect,
                    Array.Empty<double>(),
                    RotationDegrees: 0,
                    Fill: new LinearGradientShapeFill(
                        new[] { new GradientStop(0.0, ArgbColor.Black), new GradientStop(1.0, ArgbColor.White) }, 45.0),
                    Outline: null,
                    TextLines: Array.Empty<ShapeTextLine>()),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void 塗りつぶし無しの図形は枠線のみ描画される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ShapeCommand(
                    new RectPt(10, 10, 60, 20),
                    ShapePresetType.Rect,
                    Array.Empty<double>(),
                    RotationDegrees: 0,
                    Fill: null,
                    Outline: new ShapeOutline(ArgbColor.Black, 1.0),
                    TextLines: Array.Empty<ShapeTextLine>()),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void 図形内テキストはPDFに描画される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var font = new FontStyle("Calibri", 10.0, false, false, UnderlineStyle.None, false, ArgbColor.Black);
            var layout = Layout(commands: new DrawCommand[]
            {
                new ShapeCommand(
                    new RectPt(10, 10, 80, 30),
                    ShapePresetType.Rect,
                    Array.Empty<double>(),
                    RotationDegrees: 0,
                    Fill: null,
                    Outline: null,
                    TextLines: new[] { new ShapeTextLine(new PointPt(15, 25), "OK", font, TextAnchor.Left) }),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
            var content = Encoding.Latin1.GetString(output.ToArray());
            Assert.Contains("OK", content);
        }

        [Fact]
        public void 接続線はPDFに描画される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ConnectorCommand(
                    new RectPt(10, 10, 60, 20),
                    ConnectorPresetType.Bent2Segment,
                    RotationDegrees: 0,
                    FlipHorizontal: false,
                    FlipVertical: false,
                    Outline: new ShapeOutline(ArgbColor.Black, 1.0)),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void 枠線が無い接続線も既定の黒い実線で描画される()
        {
            // ConnectorCommand.Outlineがnullの場合、DrawConnectorのDefaultConnectorOutlineに
            // フォールバックする経路(design.md参照)が例外にならないことの回帰テスト。
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ConnectorCommand(
                    new RectPt(10, 10, 60, 20),
                    ConnectorPresetType.Straight,
                    RotationDegrees: 0,
                    FlipHorizontal: false,
                    FlipVertical: false,
                    Outline: null),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
            Assert.Equal("%PDF", Encoding.ASCII.GetString(output.ToArray(), 0, 4));
        }

        [Theory]
        [InlineData(ConnectorPresetType.Straight)]
        [InlineData(ConnectorPresetType.Bent2Segment)]
        [InlineData(ConnectorPresetType.Bent3Segment)]
        [InlineData(ConnectorPresetType.Curved2Segment)]
        [InlineData(ConnectorPresetType.Curved3Segment)]
        public void 全接続線プリセットが例外なく描画できる(ConnectorPresetType preset)
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var layout = Layout(commands: new DrawCommand[]
            {
                new ConnectorCommand(
                    new RectPt(10, 10, 60, 40),
                    preset,
                    RotationDegrees: 15,
                    FlipHorizontal: true,
                    FlipVertical: false,
                    Outline: new ShapeOutline(ArgbColor.Black, 1.0)),
            });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        [Fact]
        public void グループ内の図形と接続線がまとめてPDFに描画される()
        {
            // DrawGroupが未配線/誤配線だった場合に検出できる回帰テスト
            // (DrawSingleCommandへのリファクタ・GroupCommandの再帰描画の検証)。
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var childShape = new ShapeCommand(
                new RectPt(15, 15, 20, 20),
                ShapePresetType.Ellipse,
                Array.Empty<double>(),
                RotationDegrees: 0,
                Fill: new SolidShapeFill(ArgbColor.Black),
                Outline: null,
                TextLines: Array.Empty<ShapeTextLine>());
            var childConnector = new ConnectorCommand(
                new RectPt(40, 15, 20, 20),
                ConnectorPresetType.Curved3Segment,
                RotationDegrees: 0,
                FlipHorizontal: false,
                FlipVertical: true,
                Outline: new ShapeOutline(ArgbColor.Black, 1.0));

            var group = new GroupCommand(
                new PointPt(40, 30),
                RotationDegrees: 0,
                Children: new DrawCommand[] { childShape, childConnector });

            var layout = Layout(commands: new DrawCommand[] { group });

            renderer.Render(layout, output);

            var bytes = output.ToArray();
            Assert.True(bytes.Length > 0);
            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        }

        [Fact]
        public void 回転したグループとさらに回転した子図形が例外なく描画される()
        {
            // グループの回転と子要素個別の回転はcanvasの変換行列スタックで合成される
            // (design.md参照)。ピクセル単位の合成の正しさは自動テストでは検証しづらいため、
            // ここでは「例外にならず妥当なサイズのPDFが生成される」ことのみを確認する。
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var rotatedChild = new ShapeCommand(
                new RectPt(15, 15, 20, 20),
                ShapePresetType.Rect,
                Array.Empty<double>(),
                RotationDegrees: 20,
                Fill: new SolidShapeFill(ArgbColor.Black),
                Outline: null,
                TextLines: Array.Empty<ShapeTextLine>());
            var plainChild = new ShapeCommand(
                new RectPt(40, 15, 20, 20),
                ShapePresetType.Rect,
                Array.Empty<double>(),
                RotationDegrees: 0,
                Fill: new SolidShapeFill(ArgbColor.Black),
                Outline: null,
                TextLines: Array.Empty<ShapeTextLine>());

            var group = new GroupCommand(
                new PointPt(40, 30),
                RotationDegrees: 30,
                Children: new DrawCommand[] { rotatedChild, plainChild });

            var layout = Layout(commands: new DrawCommand[]
            {
                group,
                new FillRectCommand(new RectPt(100, 100, 30, 30), ArgbColor.Black),
            });

            renderer.Render(layout, output);

            var bytes = output.ToArray();
            Assert.True(bytes.Length > 0);
            Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        }

        [Fact]
        public void 入れ子のグループも例外なく描画される()
        {
            var renderer = new SkiaPdfRenderer(_metrics);
            using var output = new MemoryStream();

            var innermost = new ShapeCommand(
                new RectPt(12, 12, 5, 5),
                ShapePresetType.Rect,
                Array.Empty<double>(),
                RotationDegrees: 0,
                Fill: new SolidShapeFill(ArgbColor.Black),
                Outline: null,
                TextLines: Array.Empty<ShapeTextLine>());
            var nestedGroup = new GroupCommand(new PointPt(15, 15), RotationDegrees: 10, Children: new DrawCommand[] { innermost });
            var outerGroup = new GroupCommand(new PointPt(15, 15), RotationDegrees: 0, Children: new DrawCommand[] { nestedGroup });

            var layout = Layout(commands: new DrawCommand[] { outerGroup });

            renderer.Render(layout, output);

            Assert.True(output.Length > 0);
        }

        /// <summary>1x1のPNG(最小の有効なPNGバイト列)。</summary>
        private static byte[] TinyPng() => Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

        /// <summary>
        /// IHDRだけが指定サイズを宣言し、IDATは実際のスキャンラインと対応しない最小限のPNGを作る。
        /// <c>SKBitmap.DecodeBounds</c>がIHDRのみで寸法を報告し、IDATの整合性を検証しないことを
        /// 利用して、実データを展開せずに寸法チェックの安全性を検証するためのテスト専用ヘルパー。
        /// </summary>
        private static byte[] BuildPngWithDeclaredSize(int width, int height)
        {
            using var stream = new MemoryStream();
            stream.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, 0, 8);

            var ihdr = new byte[13];
            WriteUInt32BigEndian(ihdr, 0, (uint)width);
            WriteUInt32BigEndian(ihdr, 4, (uint)height);
            ihdr[8] = 8;
            ihdr[9] = 2;
            WriteChunk(stream, "IHDR", ihdr);

            using (var zlib = new MemoryStream())
            {
                zlib.WriteByte(0x78);
                zlib.WriteByte(0x9C);
                using (var deflate = new System.IO.Compression.DeflateStream(
                    zlib, System.IO.Compression.CompressionLevel.Fastest, leaveOpen: true))
                {
                    deflate.WriteByte(0);
                }

                zlib.Write(new byte[] { 0, 0, 0, 1 }, 0, 4); // Adler-32は検証対象外のため厳密でなくてよい
                WriteChunk(stream, "IDAT", zlib.ToArray());
            }

            WriteChunk(stream, "IEND", Array.Empty<byte>());
            return stream.ToArray();
        }

        private static void WriteChunk(Stream stream, string type, byte[] data)
        {
            var length = new byte[4];
            WriteUInt32BigEndian(length, 0, (uint)data.Length);
            stream.Write(length, 0, length.Length);

            var typeBytes = Encoding.ASCII.GetBytes(type);
            stream.Write(typeBytes, 0, typeBytes.Length);
            stream.Write(data, 0, data.Length);

            var crcInput = new byte[typeBytes.Length + data.Length];
            Buffer.BlockCopy(typeBytes, 0, crcInput, 0, typeBytes.Length);
            Buffer.BlockCopy(data, 0, crcInput, typeBytes.Length, data.Length);

            var crc = new byte[4];
            WriteUInt32BigEndian(crc, 0, Crc32(crcInput));
            stream.Write(crc, 0, crc.Length);
        }

        private static void WriteUInt32BigEndian(byte[] buffer, int offset, uint value)
        {
            buffer[offset] = (byte)(value >> 24);
            buffer[offset + 1] = (byte)(value >> 16);
            buffer[offset + 2] = (byte)(value >> 8);
            buffer[offset + 3] = (byte)value;
        }

        private static readonly uint[] CrcTable = BuildCrcTable();

        private static uint[] BuildCrcTable()
        {
            var table = new uint[256];
            for (uint n = 0; n < 256; n++)
            {
                var c = n;
                for (var k = 0; k < 8; k++)
                {
                    c = (c & 1) != 0 ? 0xEDB88320 ^ (c >> 1) : c >> 1;
                }

                table[n] = c;
            }

            return table;
        }

        private static uint Crc32(byte[] data)
        {
            var crc = 0xFFFFFFFFu;
            foreach (var value in data)
            {
                crc = CrcTable[(crc ^ value) & 0xFF] ^ (crc >> 8);
            }

            return crc ^ 0xFFFFFFFFu;
        }

        /// <summary>Rendering が解釈できない描画命令(異常系のテスト用)。</summary>
        private sealed record UnknownCommand : DrawCommand;
    }
}
