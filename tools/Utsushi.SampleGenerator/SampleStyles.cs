using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Utsushi.SampleGenerator;

/// <summary>
/// 帳票サンプルが使う書式(スタイルシート)を組み立てる。
/// </summary>
/// <remarks>
/// <see cref="Style"/> の各値が <c>cellXfs</c> の索引に対応する。
/// </remarks>
internal static class SampleStyles
{
    /// <summary>cellXfs の索引。</summary>
    internal static class Style
    {
        public const uint Default = 0;
        public const uint Title = 1;
        public const uint LabelRight = 2;
        public const uint Value = 3;
        public const uint TableHeader = 4;
        public const uint TableText = 5;
        public const uint TableNumber = 6;
        public const uint TableCurrency = 7;
        public const uint DateValue = 8;
        public const uint TotalCurrency = 9;
        public const uint Note = 10;
        public const uint CustomerName = 11;
        public const uint SectionLabel = 12;
        public const uint TableTextCentered = 13;
    }

    /// <summary>数値書式ID(組み込みと衝突しない164以降を使う)。</summary>
    private const uint CurrencyFormatId = 164;
    private const uint JapaneseDateFormatId = 165;

    /// <summary>サンプルで使う標準フォント名。</summary>
    public const string FontName = "MS PGothic";

    public static Stylesheet Create()
    {
        var numberingFormats = new NumberingFormats(
            new NumberingFormat { NumberFormatId = CurrencyFormatId, FormatCode = "\"¥\"#,##0" },
            new NumberingFormat { NumberFormatId = JapaneseDateFormatId, FormatCode = "yyyy\"年\"m\"月\"d\"日\"" })
        {
            Count = 2U,
        };

        var fonts = new Fonts(
            Font(11, bold: false),                 // 0: 標準
            Font(18, bold: true),                  // 1: タイトル
            Font(11, bold: true),                  // 2: 見出し
            Font(9, bold: false),                  // 3: 注記
            Font(14, bold: true))                  // 4: 合計
        {
            Count = 5U,
        };

        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),      // 0: 必須
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),   // 1: 必須
            new Fill(new PatternFill                                              // 2: 表見出しの背景
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = "FFDCE6F1" },
                BackgroundColor = new BackgroundColor { Indexed = 64U },
            }))
        {
            Count = 3U,
        };

        var borders = new Borders(
            new Border(                                                           // 0: 罫線なし(必須)
                new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),
            BoxBorder(BorderStyleValues.Thin),                                    // 1: 細線の box
            new Border(                                                           // 2: 下線のみ
                new LeftBorder(),
                new RightBorder(),
                new TopBorder(),
                new BottomBorder { Style = BorderStyleValues.Thin, Color = new Color { Rgb = "FF000000" } },
                new DiagonalBorder()),
            BoxBorder(BorderStyleValues.Medium))                                  // 3: 太線の box
        {
            Count = 4U,
        };

        var cellStyleFormats = new CellStyleFormats(
            new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U })
        {
            Count = 1U,
        };

        var cellFormats = new CellFormats(
            Xf(0, 0, 0, 0),                                                                        // 0: Default
            Xf(1, 0, 0, 0, HorizontalAlignmentValues.Center, VerticalAlignmentValues.Center),       // 1: Title
            Xf(0, 0, 0, 0, HorizontalAlignmentValues.Right, VerticalAlignmentValues.Center),        // 2: LabelRight
            Xf(0, 0, 0, 0, HorizontalAlignmentValues.Left, VerticalAlignmentValues.Center),         // 3: Value
            Xf(2, 2, 1, 0, HorizontalAlignmentValues.Center, VerticalAlignmentValues.Center),       // 4: TableHeader
            Xf(0, 0, 1, 0, HorizontalAlignmentValues.Left, VerticalAlignmentValues.Center),         // 5: TableText
            Xf(0, 0, 1, 3, HorizontalAlignmentValues.Right, VerticalAlignmentValues.Center),        // 6: TableNumber
            Xf(0, 0, 1, (int)CurrencyFormatId, HorizontalAlignmentValues.Right, VerticalAlignmentValues.Center), // 7
            Xf(0, 0, 0, (int)JapaneseDateFormatId, HorizontalAlignmentValues.Left, VerticalAlignmentValues.Center), // 8
            Xf(4, 0, 3, (int)CurrencyFormatId, HorizontalAlignmentValues.Right, VerticalAlignmentValues.Center),    // 9
            Xf(3, 0, 0, 0, HorizontalAlignmentValues.Left, VerticalAlignmentValues.Top, wrap: true),  // 10: Note
            Xf(2, 0, 2, 0, HorizontalAlignmentValues.Left, VerticalAlignmentValues.Bottom),           // 11: CustomerName
            Xf(2, 0, 0, 0, HorizontalAlignmentValues.Left, VerticalAlignmentValues.Center),           // 12: SectionLabel
            Xf(0, 0, 1, 0, HorizontalAlignmentValues.Center, VerticalAlignmentValues.Center))         // 13: TableTextCentered
        {
            Count = 14U,
        };

        return new Stylesheet(numberingFormats, fonts, fills, borders, cellStyleFormats, cellFormats);
    }

    private static Font Font(double size, bool bold)
    {
        var font = new Font(
            new FontSize { Val = size },
            new Color { Rgb = "FF000000" },
            new FontName { Val = FontName },
            new FontFamilyNumbering { Val = 3 });

        if (bold)
        {
            font.InsertAt(new Bold(), 0);
        }

        return font;
    }

    private static Border BoxBorder(BorderStyleValues style)
    {
        var color = new Color { Rgb = "FF000000" };
        return new Border(
            new LeftBorder(color.CloneNode(true)) { Style = style },
            new RightBorder(color.CloneNode(true)) { Style = style },
            new TopBorder(color.CloneNode(true)) { Style = style },
            new BottomBorder(color.CloneNode(true)) { Style = style },
            new DiagonalBorder());
    }

    private static CellFormat Xf(
        int fontId,
        int fillId,
        int borderId,
        int numberFormatId,
        HorizontalAlignmentValues? horizontal = null,
        VerticalAlignmentValues? vertical = null,
        bool wrap = false)
    {
        var format = new CellFormat
        {
            FontId = (uint)fontId,
            FillId = (uint)fillId,
            BorderId = (uint)borderId,
            NumberFormatId = (uint)numberFormatId,
            ApplyFont = true,
            ApplyFill = true,
            ApplyBorder = true,
            ApplyNumberFormat = numberFormatId != 0,
        };

        if (horizontal is null && vertical is null && !wrap)
        {
            return format;
        }

        var alignment = new Alignment();
        if (horizontal is { } h)
        {
            alignment.Horizontal = h;
        }

        if (vertical is { } v)
        {
            alignment.Vertical = v;
        }

        if (wrap)
        {
            alignment.WrapText = true;
        }

        format.Alignment = alignment;
        format.ApplyAlignment = true;
        return format;
    }
}
