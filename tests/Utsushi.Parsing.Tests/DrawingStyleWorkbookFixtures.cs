using System;
using System.Globalization;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// テーマの色・スタイル参照(要件10.15, 10.16)、反転(要件10.17)、矢印(要件10.18)の読み取りテスト用に、
    /// <c>drawing1.xml</c>・<c>theme1.xml</c> を XML 文字列のまま書き込んだ最小の .xlsx を組み立てるヘルパー。
    /// </summary>
    /// <remarks>
    /// Excel が実際に保存する XML(<c>xdr:style</c> 等)をそのまま再現したいため、SDK の型を組み立てずに文字列で書く。
    /// </remarks>
    internal static class DrawingStyleWorkbookFixtures
    {
        public const string DrawingNamespaces =
            "xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
            + "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"";

        /// <summary>Excel が挿入した図形に付ける既定のスタイル(要件10.16)。</summary>
        public const string ExcelDefaultShapeStyle =
            "<xdr:style>"
            + "<a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef>"
            + "<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>"
            + "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>"
            + "<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef>"
            + "</xdr:style>";

        /// <summary>Excel が挿入した接続線に付ける既定のスタイル(要件10.16)。</summary>
        public const string ExcelDefaultConnectorStyle =
            "<xdr:style>"
            + "<a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef>"
            + "<a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>"
            + "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>"
            + "<a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef>"
            + "</xdr:style>";

        /// <summary>
        /// 指定したアンカー群(<c>xdr:oneCellAnchor</c> 等の XML)を持つ .xlsx を一時ファイルとして作成する。
        /// <paramref name="themeXml"/> が null ならテーマパートを作らない。呼び出し側で削除すること。
        /// </summary>
        public static string CreateWorkbook(string anchorsXml, string? themeXml = null)
        {
            var path = Path.Combine(Path.GetTempPath(), "utsushi-style-test-" + Guid.NewGuid().ToString("N") + ".xlsx");

            using (var document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                if (themeXml is not null)
                {
                    var themePart = workbookPart.AddNewPart<ThemePart>();
                    Feed(themePart, themeXml);
                }

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "テストシート",
                });

                var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                Feed(drawingsPart, "<xdr:wsDr " + DrawingNamespaces + ">" + anchorsXml + "</xdr:wsDr>");

                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            return path;
        }

        /// <summary>
        /// テーマ(<c>a:theme</c>)の XML。配色は Office 既定のテーマに合わせ、<paramref name="accent1"/> と
        /// <paramref name="dark1"/> だけ差し替えられる。<paramref name="lineStyleWidthsEmu"/> は <c>lnStyleLst</c> の各
        /// <c>a:ln/@w</c>(null の要素は <c>@w</c> を書かない)。
        /// </summary>
        public static string ThemeXml(
            string accent1 = "4472C4",
            string dark1 = "000000",
            params long?[] lineStyleWidthsEmu)
        {
            var lineStyles = new StringBuilder();
            foreach (var width in lineStyleWidthsEmu)
            {
                lineStyles.Append(width is { } w ? "<a:ln w=\"" + w.ToString(CultureInfo.InvariantCulture) + "\">" : "<a:ln>");
                lineStyles.Append("<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>");
            }

            return "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">"
                + "<a:themeElements>"
                + "<a:clrScheme name=\"Office\">"
                + "<a:dk1><a:srgbClr val=\"" + dark1 + "\"/></a:dk1>"
                + "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>"
                + "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>"
                + "<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>"
                + "<a:accent1><a:srgbClr val=\"" + accent1 + "\"/></a:accent1>"
                + "<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>"
                + "<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>"
                + "<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>"
                + "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>"
                + "<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>"
                + "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>"
                + "<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>"
                + "</a:clrScheme>"
                + "<a:fontScheme name=\"Office\">"
                + "<a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>"
                + "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>"
                + "</a:fontScheme>"
                + "<a:fmtScheme name=\"Office\">"
                + "<a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
                + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
                + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>"
                + "<a:lnStyleLst>" + lineStyles + "</a:lnStyleLst>"
                + "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>"
                + "<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
                + "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
                + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
                + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst>"
                + "</a:fmtScheme>"
                + "</a:themeElements>"
                + "</a:theme>";
        }

        /// <summary>Office 既定のテーマ(<c>lnStyleLst</c> は 0.5pt / 1pt / 1.5pt)。</summary>
        public static string OfficeThemeXml() => ThemeXml(lineStyleWidthsEmu: new long?[] { 6350, 12700, 19050 });

        /// <summary>
        /// <c>xdr:sp</c> を1つ持つ oneCellAnchor(B3 起点・900000EMU 四方)。<paramref name="xfrmAttributes"/> は
        /// <c>a:xfrm</c> に付ける属性(例: <c>flipH="1"</c>)、<paramref name="spPrExtra"/> は <c>a:prstGeom</c> の後ろに
        /// 置く塗りつぶし・線の XML、<paramref name="style"/> は <c>xdr:style</c>、<paramref name="textBody"/> は <c>xdr:txBody</c>。
        /// </summary>
        public static string ShapeAnchor(
            string spPrExtra = "",
            string style = "",
            string textBody = "",
            string xfrmAttributes = "",
            string preset = "rect",
            uint id = 2U) =>
            "<xdr:oneCellAnchor>"
            + "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            + "<xdr:ext cx=\"900000\" cy=\"900000\"/>"
            + ShapeElement(spPrExtra, style, textBody, xfrmAttributes, preset, id, offX: 0, offY: 0, cx: 900000, cy: 900000)
            + "<xdr:clientData/>"
            + "</xdr:oneCellAnchor>";

        /// <summary><c>xdr:sp</c> 要素本体(グループの子としても使う)。</summary>
        public static string ShapeElement(
            string spPrExtra = "",
            string style = "",
            string textBody = "",
            string xfrmAttributes = "",
            string preset = "rect",
            uint id = 2U,
            long offX = 0,
            long offY = 0,
            long cx = 900000,
            long cy = 900000) =>
            "<xdr:sp macro=\"\" textlink=\"\">"
            + "<xdr:nvSpPr><xdr:cNvPr id=\"" + id.ToString(CultureInfo.InvariantCulture) + "\" name=\"Shape\"/><xdr:cNvSpPr/></xdr:nvSpPr>"
            + "<xdr:spPr>"
            + Xfrm(xfrmAttributes, offX, offY, cx, cy)
            + "<a:prstGeom prst=\"" + preset + "\"><a:avLst/></a:prstGeom>"
            + spPrExtra
            + "</xdr:spPr>"
            + style
            + textBody
            + "</xdr:sp>";

        /// <summary><c>xdr:cxnSp</c> を1つ持つ oneCellAnchor。</summary>
        public static string ConnectorAnchor(string spPrExtra = "", string style = "", uint id = 3U) =>
            "<xdr:oneCellAnchor>"
            + "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            + "<xdr:ext cx=\"900000\" cy=\"900000\"/>"
            + ConnectorElement(spPrExtra, style, id)
            + "<xdr:clientData/>"
            + "</xdr:oneCellAnchor>";

        /// <summary><c>xdr:cxnSp</c> 要素本体(グループの子としても使う)。</summary>
        public static string ConnectorElement(string spPrExtra = "", string style = "", uint id = 3U) =>
            "<xdr:cxnSp macro=\"\">"
            + "<xdr:nvCxnSpPr><xdr:cNvPr id=\"" + id.ToString(CultureInfo.InvariantCulture) + "\" name=\"Connector\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>"
            + "<xdr:spPr>"
            + Xfrm(string.Empty, 0, 0, 900000, 900000)
            + "<a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom>"
            + spPrExtra
            + "</xdr:spPr>"
            + style
            + "</xdr:cxnSp>";

        /// <summary>
        /// トップレベルのグループ(<c>xdr:grpSp</c>)を1つ持つ oneCellAnchor。子座標空間は 1800000EMU 四方。
        /// </summary>
        public static string GroupAnchor(string childrenXml, string xfrmAttributes = "", uint id = 10U) =>
            "<xdr:oneCellAnchor>"
            + "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            + "<xdr:ext cx=\"1800000\" cy=\"1800000\"/>"
            + GroupElement(childrenXml, xfrmAttributes, id, 0, 0, 1800000, 1800000)
            + "<xdr:clientData/>"
            + "</xdr:oneCellAnchor>";

        /// <summary><c>xdr:grpSp</c> 要素本体(入れ子のグループとしても使う)。子座標空間は配置と同じ大きさ。</summary>
        public static string GroupElement(string childrenXml, string xfrmAttributes, uint id, long offX, long offY, long cx, long cy) =>
            "<xdr:grpSp>"
            + "<xdr:nvGrpSpPr><xdr:cNvPr id=\"" + id.ToString(CultureInfo.InvariantCulture) + "\" name=\"Group\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>"
            + "<xdr:grpSpPr>"
            + "<a:xfrm " + xfrmAttributes + ">"
            + "<a:off x=\"" + N(offX) + "\" y=\"" + N(offY) + "\"/><a:ext cx=\"" + N(cx) + "\" cy=\"" + N(cy) + "\"/>"
            + "<a:chOff x=\"" + N(offX) + "\" y=\"" + N(offY) + "\"/><a:chExt cx=\"" + N(cx) + "\" cy=\"" + N(cy) + "\"/>"
            + "</a:xfrm>"
            + "</xdr:grpSpPr>"
            + childrenXml
            + "</xdr:grpSp>";

        /// <summary>1段落・1ランのテキスト。<paramref name="runPropertiesInner"/> は <c>a:rPr</c> の子要素。</summary>
        public static string TextBody(string text, string? runPropertiesInner = null) =>
            "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>"
            + (runPropertiesInner is null ? string.Empty : "<a:rPr lang=\"ja-JP\" sz=\"1100\">" + runPropertiesInner + "</a:rPr>")
            + "<a:t>" + text + "</a:t></a:r></a:p></xdr:txBody>";

        private static string Xfrm(string attributes, long offX, long offY, long cx, long cy) =>
            "<a:xfrm " + attributes + "><a:off x=\"" + N(offX) + "\" y=\"" + N(offY) + "\"/>"
            + "<a:ext cx=\"" + N(cx) + "\" cy=\"" + N(cy) + "\"/></a:xfrm>";

        private static string N(long value) => value.ToString(CultureInfo.InvariantCulture);

        private static void Feed(OpenXmlPart part, string xml)
        {
            using var stream = new MemoryStream(new UTF8Encoding(false).GetBytes(xml));
            part.FeedData(stream);
        }
    }
}
