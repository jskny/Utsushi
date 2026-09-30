using System;
using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;
using static Utsushi.Parsing.Tests.DrawingStyleWorkbookFixtures;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 図形・接続線のテーマの色とスタイル参照(要件10.15, 10.16)、反転(要件10.17)、矢印(要件10.18)の読み取りの検証。
    /// </summary>
    public sealed class ShapeStyleAndFlipReadingTests
    {
        private static readonly ArgbColor Accent1 = Hex("4472C4");

        /// <summary>accent1 に shade 50% をかけた色。Excel の既定の図形の枠線の色。</summary>
        private static readonly ArgbColor Accent1Shade50 = Hex("2F528F");

        private readonly OpenXmlWorkbookReader _reader = new();

        private static ArgbColor Hex(string hex)
        {
            Assert.True(ArgbColor.TryParseHex(hex, out var color));
            return color;
        }

        private SheetModel Read(string anchorsXml, string? themeXml = null)
        {
            var path = CreateWorkbook(anchorsXml, themeXml);
            try
            {
                return Assert.Single(_reader.ReadFile(path).Sheets);
            }
            finally
            {
                File.Delete(path);
            }
        }

        private ShapeModel ReadShape(string anchorXml, string? themeXml = null) =>
            Assert.Single(Read(anchorXml, themeXml).DrawingObjects.OfType<ShapeModel>());

        private ConnectorModel ReadConnector(string anchorXml, string? themeXml = null) =>
            Assert.Single(Read(anchorXml, themeXml).DrawingObjects.OfType<ConnectorModel>());

        // -- 要件10.16: 図形のスタイル参照 -------------------------------------------------

        [Fact]
        public void Excelの既定の図形スタイルだけを持つ図形は塗りと枠線と文字色がスタイルで決まる()
        {
            var shape = ReadShape(
                ShapeAnchor(style: ExcelDefaultShapeStyle, textBody: TextBody("見積")),
                OfficeThemeXml());

            // fillRef idx=1 accent1 → 塗り、lnRef idx=2 accent1 shade 50% → 枠線(太さはテーマの2番目の線 = 1pt)、
            // fontRef lt1 → 文字色(白)。
            Assert.Equal(Accent1, Assert.IsType<SolidShapeFill>(shape.Fill).Color);
            Assert.NotNull(shape.Outline);
            Assert.Equal(Accent1Shade50, shape.Outline!.Color);
            Assert.Equal(1.0, shape.Outline.WidthPt, 6);
            Assert.Null(shape.Outline.HeadEnd);
            Assert.Null(shape.Outline.TailEnd);

            var run = Assert.Single(Assert.Single(shape.Text!.Paragraphs).Runs);
            Assert.Equal(ArgbColor.White, run.Font.Color);
        }

        [Fact]
        public void テーマの無いブックでは既定の配色と既定の線の太さでスタイルを解決する()
        {
            var shape = ReadShape(ShapeAnchor(style: ExcelDefaultShapeStyle, textBody: TextBody("見積")), themeXml: null);

            Assert.Equal(Accent1, Assert.IsType<SolidShapeFill>(shape.Fill).Color);
            Assert.Equal(Accent1Shade50, shape.Outline!.Color);

            // lnRef idx=2 は Office 既定テーマの lnStyleLst の2番目(1pt)。
            Assert.Equal(1.0, shape.Outline.WidthPt, 6);
        }

        [Fact]
        public void スタイルの色はブックのテーマの配色から求める()
        {
            var shape = ReadShape(
                ShapeAnchor(style: ExcelDefaultShapeStyle),
                ThemeXml(accent1: "112233", lineStyleWidthsEmu: new long?[] { 6350, 25400, 38100 }));

            Assert.Equal(Hex("112233"), Assert.IsType<SolidShapeFill>(shape.Fill).Color);
            Assert.Equal(2.0, shape.Outline!.WidthPt, 6);
        }

        [Fact]
        public void spPrの塗りと線の指定はスタイルより優先する()
        {
            var shape = ReadShape(
                ShapeAnchor(
                    spPrExtra: "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>"
                        + "<a:ln w=\"38100\"><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln>",
                    style: ExcelDefaultShapeStyle,
                    textBody: TextBody("見積", "<a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill>")),
                OfficeThemeXml());

            Assert.Equal(Hex("FF0000"), Assert.IsType<SolidShapeFill>(shape.Fill).Color);
            Assert.Equal(Hex("00FF00"), shape.Outline!.Color);
            Assert.Equal(3.0, shape.Outline.WidthPt, 6);
            Assert.Equal(Hex("0000FF"), Assert.Single(Assert.Single(shape.Text!.Paragraphs).Runs).Font.Color);
        }

        [Fact]
        public void 線の太さだけを指定した場合は色をスタイルから太さをspPrから求める()
        {
            var shape = ReadShape(
                ShapeAnchor(spPrExtra: "<a:ln w=\"25400\"/>", style: ExcelDefaultShapeStyle),
                OfficeThemeXml());

            Assert.Equal(Accent1Shade50, shape.Outline!.Color);
            Assert.Equal(2.0, shape.Outline.WidthPt, 6);
        }

        [Fact]
        public void 線の色だけを指定した場合は太さをスタイルの線の書式から求める()
        {
            var shape = ReadShape(
                ShapeAnchor(
                    spPrExtra: "<a:ln><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln>",
                    style: ExcelDefaultShapeStyle),
                OfficeThemeXml());

            Assert.Equal(Hex("00FF00"), shape.Outline!.Color);
            Assert.Equal(1.0, shape.Outline.WidthPt, 6); // lnRef idx=2 → 12700EMU
        }

        [Fact]
        public void スタイルも太さも無い線は1ptになる()
        {
            var shape = ReadShape(ShapeAnchor(spPrExtra: "<a:ln><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln>"));

            Assert.Equal(1.0, shape.Outline!.WidthPt, 6);
        }

        [Fact]
        public void spPrのnoFillはスタイルの塗りより優先して塗りなしになる()
        {
            var shape = ReadShape(ShapeAnchor(spPrExtra: "<a:noFill/>", style: ExcelDefaultShapeStyle), OfficeThemeXml());

            Assert.Null(shape.Fill);
            Assert.NotNull(shape.Outline);
        }

        [Fact]
        public void 線のnoFillはスタイルの線より優先して線なしになる()
        {
            var shape = ReadShape(
                ShapeAnchor(spPrExtra: "<a:ln w=\"12700\"><a:noFill/></a:ln>", style: ExcelDefaultShapeStyle),
                OfficeThemeXml());

            Assert.NotNull(shape.Fill);
            Assert.Null(shape.Outline);
        }

        [Fact]
        public void fillRefとlnRefの番号が0なら塗りなしと線なしになる()
        {
            const string style =
                "<xdr:style>"
                + "<a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef>"
                + "<a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>"
                + "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>"
                + "<a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef>"
                + "</xdr:style>";

            var shape = ReadShape(ShapeAnchor(style: style, textBody: TextBody("見積")), OfficeThemeXml());

            Assert.Null(shape.Fill);
            Assert.Null(shape.Outline);
            Assert.Equal(ArgbColor.Black, Assert.Single(Assert.Single(shape.Text!.Paragraphs).Runs).Font.Color);
        }

        [Fact]
        public void lnRefの番号が0でもspPrに線の色があれば1ptの線になる()
        {
            const string style =
                "<xdr:style>"
                + "<a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef>"
                + "<a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>"
                + "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>"
                + "<a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef>"
                + "</xdr:style>";

            var shape = ReadShape(
                ShapeAnchor(spPrExtra: "<a:ln><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln>", style: style),
                OfficeThemeXml());

            Assert.Equal(Hex("00FF00"), shape.Outline!.Color);
            Assert.Equal(1.0, shape.Outline.WidthPt, 6);
        }

        [Theory]
        [InlineData("<a:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></a:blipFill>")]
        [InlineData("<a:pattFill prst=\"pct5\"><a:fgClr><a:srgbClr val=\"FF0000\"/></a:fgClr><a:bgClr><a:srgbClr val=\"FFFFFF\"/></a:bgClr></a:pattFill>")]
        [InlineData("<a:grpFill/>")]
        public void 未対応の塗りはスタイルの塗りで代用しない(string fillXml)
        {
            var shape = ReadShape(ShapeAnchor(spPrExtra: fillXml, style: ExcelDefaultShapeStyle), OfficeThemeXml());

            Assert.Null(shape.Fill);
        }

        [Theory]
        [InlineData("<a:ln w=\"12700\"><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst></a:gradFill></a:ln>")]
        [InlineData("<a:ln w=\"12700\"><a:pattFill prst=\"pct5\"><a:fgClr><a:srgbClr val=\"FF0000\"/></a:fgClr><a:bgClr><a:srgbClr val=\"FFFFFF\"/></a:bgClr></a:pattFill></a:ln>")]
        public void 未対応の線はスタイルの線で代用しない(string lineXml)
        {
            var shape = ReadShape(ShapeAnchor(spPrExtra: lineXml, style: ExcelDefaultShapeStyle), OfficeThemeXml());

            Assert.Null(shape.Outline);
        }

        [Fact]
        public void spPrの色が解決できない場合はスタイルで代用せず塗りなしになる()
        {
            var shape = ReadShape(
                ShapeAnchor(spPrExtra: "<a:solidFill><a:prstClr val=\"red\"/></a:solidFill>", style: ExcelDefaultShapeStyle),
                OfficeThemeXml());

            Assert.Null(shape.Fill);
        }

        // -- 要件10.15: テーマの色 -------------------------------------------------

        [Fact]
        public void spPrと文字のschemeClrをテーマの配色と修飾で解決する()
        {
            var shape = ReadShape(
                ShapeAnchor(
                    spPrExtra: "<a:solidFill><a:schemeClr val=\"accent1\"><a:lumMod val=\"75000\"/></a:schemeClr></a:solidFill>"
                        + "<a:ln w=\"12700\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:ln>",
                    textBody: TextBody("見積", "<a:solidFill><a:schemeClr val=\"accent2\"/></a:solidFill>")),
                OfficeThemeXml());

            Assert.Equal(Hex("2F5597"), Assert.IsType<SolidShapeFill>(shape.Fill).Color);
            Assert.Equal(ArgbColor.Black, shape.Outline!.Color);
            Assert.Equal(Hex("ED7D31"), Assert.Single(Assert.Single(shape.Text!.Paragraphs).Runs).Font.Color);
        }

        [Fact]
        public void グラデーションの分岐点のschemeClrも解決する()
        {
            var shape = ReadShape(
                ShapeAnchor(spPrExtra:
                    "<a:gradFill><a:gsLst>"
                    + "<a:gs pos=\"0\"><a:schemeClr val=\"accent1\"/></a:gs>"
                    + "<a:gs pos=\"100000\"><a:schemeClr val=\"accent1\"><a:tint val=\"50000\"/></a:schemeClr></a:gs>"
                    + "</a:gsLst><a:lin ang=\"0\" scaled=\"1\"/></a:gradFill>"),
                OfficeThemeXml());

            var fill = Assert.IsType<LinearGradientShapeFill>(shape.Fill);
            Assert.Equal(Accent1, fill.Stops[0].Color);
            Assert.Equal(Hex("C0C9E4"), fill.Stops[1].Color);
        }

        // -- 要件10.16: 接続線のスタイル参照 -------------------------------------------------

        [Fact]
        public void 接続線の既定スタイルのlnRefから線の色と太さを求める()
        {
            var connector = ReadConnector(ConnectorAnchor(style: ExcelDefaultConnectorStyle), OfficeThemeXml());

            Assert.NotNull(connector.Outline);
            Assert.Equal(Accent1, connector.Outline!.Color);
            Assert.Equal(0.5, connector.Outline.WidthPt, 6); // lnRef idx=1 → 6350EMU
        }

        [Fact]
        public void 接続線のspPrの線はスタイルより優先する()
        {
            var connector = ReadConnector(
                ConnectorAnchor(
                    spPrExtra: "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln>",
                    style: ExcelDefaultConnectorStyle),
                OfficeThemeXml());

            Assert.Equal(Hex("FF0000"), connector.Outline!.Color);
            Assert.Equal(2.0, connector.Outline.WidthPt, 6);
        }

        // -- 要件10.16: 接続線の線を明示的に消した場合 -------------------------------------------------
        // 接続線は Outline が null のとき Rendering が既定の黒い線(1pt)を補うため、線を明示的に消した場合は
        // 透明・0pt の線にして区別する。a:ln も xdr:style も無い場合だけが null(既定の黒い線)になる。

        /// <summary>線の番号(lnRef/@idx)が0の、線を持たない接続線のスタイル。</summary>
        private const string ConnectorStyleWithoutLine =
            "<xdr:style>"
            + "<a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef>"
            + "<a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>"
            + "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>"
            + "<a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef>"
            + "</xdr:style>";

        private static void AssertTransparentLine(ShapeOutline? outline)
        {
            Assert.NotNull(outline);
            Assert.Equal(ArgbColor.Transparent, outline!.Color);
            Assert.True(outline.Color.IsTransparent);
            Assert.Equal(0.0, outline.WidthPt);
            Assert.Null(outline.HeadEnd);
            Assert.Null(outline.TailEnd);
        }

        [Fact]
        public void 接続線の線のnoFillは透明の線になる()
        {
            var connector = ReadConnector(ConnectorAnchor(spPrExtra: "<a:ln w=\"12700\"><a:noFill/></a:ln>"));

            AssertTransparentLine(connector.Outline);
        }

        [Fact]
        public void 接続線の線のnoFillはスタイルの線より優先して透明の線になる()
        {
            var connector = ReadConnector(
                ConnectorAnchor(spPrExtra: "<a:ln w=\"12700\"><a:noFill/></a:ln>", style: ExcelDefaultConnectorStyle),
                OfficeThemeXml());

            AssertTransparentLine(connector.Outline);
        }

        [Fact]
        public void 接続線のスタイルのlnRefの番号が0なら透明の線になる()
        {
            var connector = ReadConnector(ConnectorAnchor(style: ConnectorStyleWithoutLine), OfficeThemeXml());

            AssertTransparentLine(connector.Outline);
        }

        [Fact]
        public void 接続線のスタイルの線の番号が0でもspPrに線の色があればその線になる()
        {
            var connector = ReadConnector(
                ConnectorAnchor(
                    spPrExtra: "<a:ln><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln>",
                    style: ConnectorStyleWithoutLine),
                OfficeThemeXml());

            Assert.Equal(Hex("00FF00"), connector.Outline!.Color);
            Assert.Equal(1.0, connector.Outline.WidthPt, 6);
        }

        [Fact]
        public void 接続線に線の指定もスタイルも無ければnullになり描画時に既定の線が補われる()
        {
            var connector = ReadConnector(ConnectorAnchor());

            Assert.Null(connector.Outline);
        }

        [Fact]
        public void 接続線の線に色が無くスタイルも無ければnullになる()
        {
            // 太さだけの a:ln は線を消す指定ではない。
            var connector = ReadConnector(ConnectorAnchor(spPrExtra: "<a:ln w=\"25400\"/>"));

            Assert.Null(connector.Outline);
        }

        [Fact]
        public void 接続線の色の無い線に矢印だけがあれば既定の黒い線に矢印を付ける()
        {
            // 回帰テスト(code-reviewer指摘): 以前は null になり、描画時に既定の線は補われるのに矢印だけが消えていた。
            var connector = ReadConnector(ConnectorAnchor(spPrExtra: "<a:ln w=\"25400\"><a:tailEnd type=\"triangle\"/></a:ln>"));

            Assert.NotNull(connector.Outline);
            Assert.Equal(ArgbColor.Black, connector.Outline!.Color);
            Assert.Equal(2.0, connector.Outline.WidthPt, 6);
            Assert.Null(connector.Outline.HeadEnd);
            Assert.Equal(LineEndType.Triangle, connector.Outline.TailEnd!.Type);
        }

        [Fact]
        public void 図形の線のnoFillは透明の線ではなくnullのまま()
        {
            // 図形は Outline が null なら枠線を描かない(既定の線を補うのは接続線だけ)。
            var shape = ReadShape(ShapeAnchor(spPrExtra: "<a:ln w=\"12700\"><a:noFill/></a:ln>"));

            Assert.Null(shape.Outline);
        }

        [Fact]
        public void グループ内の接続線も線を消した場合は透明の線になる()
        {
            var sheet = Read(
                GroupAnchor(
                    ConnectorElement(spPrExtra: "<a:ln w=\"12700\"><a:noFill/></a:ln>", id: 11U)
                    + ConnectorElement(style: ConnectorStyleWithoutLine, id: 12U)
                    + ConnectorElement(id: 13U)),
                OfficeThemeXml());

            var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
            var connectors = group.Children.OfType<GroupChildConnector>().ToList();
            Assert.Equal(3, connectors.Count);

            AssertTransparentLine(connectors[0].Outline);
            AssertTransparentLine(connectors[1].Outline);
            Assert.Null(connectors[2].Outline);
        }

        // -- グループ内の図形・接続線 -------------------------------------------------

        [Fact]
        public void グループ内の図形と接続線もトップレベルと同じくスタイルを解決する()
        {
            var sheet = Read(
                GroupAnchor(
                    ShapeElement(style: ExcelDefaultShapeStyle, textBody: TextBody("見積"), id: 11U)
                    + ConnectorElement(style: ExcelDefaultConnectorStyle, id: 12U)),
                OfficeThemeXml());

            var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
            var shape = Assert.Single(group.Children.OfType<GroupChildShape>());
            var connector = Assert.Single(group.Children.OfType<GroupChildConnector>());

            var topLevel = ReadShape(ShapeAnchor(style: ExcelDefaultShapeStyle, textBody: TextBody("見積")), OfficeThemeXml());
            Assert.Equal(topLevel.Fill, shape.Fill);
            Assert.Equal(topLevel.Outline, shape.Outline);
            Assert.Equal(
                Assert.Single(Assert.Single(topLevel.Text!.Paragraphs).Runs).Font.Color,
                Assert.Single(Assert.Single(shape.Text!.Paragraphs).Runs).Font.Color);

            Assert.Equal(Accent1, connector.Outline!.Color);
            Assert.Equal(0.5, connector.Outline.WidthPt, 6);
        }

        // -- 要件10.17: 反転 -------------------------------------------------

        [Theory]
        [InlineData("", false, false)]
        [InlineData("flipH=\"1\"", true, false)]
        [InlineData("flipV=\"1\"", false, true)]
        [InlineData("flipH=\"1\" flipV=\"1\"", true, true)]
        [InlineData("flipH=\"0\" flipV=\"0\"", false, false)]
        public void 図形の反転を読み取る(string xfrmAttributes, bool expectedH, bool expectedV)
        {
            var shape = ReadShape(ShapeAnchor(xfrmAttributes: xfrmAttributes));

            Assert.Equal(expectedH, shape.FlipHorizontal);
            Assert.Equal(expectedV, shape.FlipVertical);
        }

        [Fact]
        public void グループとグループ内図形と入れ子のグループの反転を読み取る()
        {
            var nested = GroupElement(
                ShapeElement(xfrmAttributes: "flipV=\"1\"", id: 14U, cx: 450000, cy: 450000),
                "flipV=\"1\"",
                13U,
                900000,
                900000,
                900000,
                900000);
            var sheet = Read(GroupAnchor(ShapeElement(xfrmAttributes: "flipH=\"1\"", id: 11U) + nested, "flipH=\"1\""));

            var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
            Assert.True(group.FlipHorizontal);
            Assert.False(group.FlipVertical);

            var child = Assert.Single(group.Children.OfType<GroupChildShape>());
            Assert.True(child.FlipHorizontal);
            Assert.False(child.FlipVertical);

            var nestedGroup = Assert.Single(group.Children.OfType<GroupChildGroup>());
            Assert.False(nestedGroup.FlipHorizontal);
            Assert.True(nestedGroup.FlipVertical);

            var nestedChild = Assert.IsType<GroupChildShape>(Assert.Single(nestedGroup.Children));
            Assert.False(nestedChild.FlipHorizontal);
            Assert.True(nestedChild.FlipVertical);
        }

        [Fact]
        public void 反転の指定が無いグループは反転しない()
        {
            var sheet = Read(GroupAnchor(ShapeElement(id: 11U)));

            var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
            Assert.False(group.FlipHorizontal);
            Assert.False(group.FlipVertical);
            var child = Assert.IsType<GroupChildShape>(Assert.Single(group.Children));
            Assert.False(child.FlipHorizontal);
            Assert.False(child.FlipVertical);
        }

        // -- 要件10.18: 矢印 -------------------------------------------------

        [Theory]
        [InlineData("triangle", LineEndType.Triangle)]
        [InlineData("stealth", LineEndType.Stealth)]
        [InlineData("arrow", LineEndType.Arrow)]
        [InlineData("oval", LineEndType.Oval)]
        [InlineData("diamond", LineEndType.Diamond)]
        public void 線の端の矢印の種類を読み取る(string type, LineEndType expected)
        {
            var connector = ReadConnector(ConnectorAnchor(
                spPrExtra: "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>"
                    + "<a:headEnd type=\"" + type + "\"/><a:tailEnd type=\"" + type + "\"/></a:ln>"));

            Assert.Equal(new LineEndStyle(expected, LineEndSize.Medium, LineEndSize.Medium), connector.Outline!.HeadEnd);
            Assert.Equal(new LineEndStyle(expected, LineEndSize.Medium, LineEndSize.Medium), connector.Outline.TailEnd);
        }

        [Theory]
        [InlineData("w=\"sm\" len=\"lg\"", LineEndSize.Small, LineEndSize.Large)]
        [InlineData("w=\"lg\" len=\"sm\"", LineEndSize.Large, LineEndSize.Small)]
        [InlineData("w=\"med\" len=\"med\"", LineEndSize.Medium, LineEndSize.Medium)]
        [InlineData("", LineEndSize.Medium, LineEndSize.Medium)]
        public void 線の端の矢印の幅と長さを読み取る(string sizeAttributes, LineEndSize expectedWidth, LineEndSize expectedLength)
        {
            var connector = ReadConnector(ConnectorAnchor(
                spPrExtra: "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>"
                    + "<a:tailEnd type=\"triangle\" " + sizeAttributes + "/></a:ln>"));

            Assert.Null(connector.Outline!.HeadEnd);
            Assert.Equal(new LineEndStyle(LineEndType.Triangle, expectedWidth, expectedLength), connector.Outline.TailEnd);
        }

        [Fact]
        public void 種類がnoneの矢印と指定の無い矢印は矢印なしになる()
        {
            var connector = ReadConnector(ConnectorAnchor(
                spPrExtra: "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>"
                    + "<a:headEnd type=\"none\" w=\"lg\" len=\"lg\"/><a:tailEnd/></a:ln>"));

            Assert.NotNull(connector.Outline);
            Assert.Null(connector.Outline!.HeadEnd);
            Assert.Null(connector.Outline.TailEnd);
        }

        [Fact]
        public void 線の色をスタイルから求める場合も矢印を読み取る()
        {
            var connector = ReadConnector(
                ConnectorAnchor(
                    spPrExtra: "<a:ln><a:tailEnd type=\"stealth\" w=\"lg\" len=\"lg\"/></a:ln>",
                    style: ExcelDefaultConnectorStyle),
                OfficeThemeXml());

            Assert.Equal(Accent1, connector.Outline!.Color);
            Assert.Equal(new LineEndStyle(LineEndType.Stealth, LineEndSize.Large, LineEndSize.Large), connector.Outline.TailEnd);
        }

        [Fact]
        public void 図形とグループ内接続線の枠線の矢印も読み取る()
        {
            const string line =
                "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>"
                + "<a:headEnd type=\"oval\" w=\"sm\" len=\"sm\"/></a:ln>";

            var shape = ReadShape(ShapeAnchor(spPrExtra: line, preset: "borderCallout1"));
            Assert.Equal(new LineEndStyle(LineEndType.Oval, LineEndSize.Small, LineEndSize.Small), shape.Outline!.HeadEnd);

            var sheet = Read(GroupAnchor(ConnectorElement(spPrExtra: line, id: 12U)));
            var group = Assert.Single(sheet.DrawingObjects.OfType<GroupShapeModel>());
            var connector = Assert.IsType<GroupChildConnector>(Assert.Single(group.Children));
            Assert.Equal(new LineEndStyle(LineEndType.Oval, LineEndSize.Small, LineEndSize.Small), connector.Outline!.HeadEnd);
        }
    }
}
