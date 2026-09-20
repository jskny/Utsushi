using Utsushi.Core;

namespace Utsushi.Parsing.Model
{
    /// <summary>水平方向の配置。</summary>
    public enum HorizontalAlignment
    {
        /// <summary>未指定。値の型に応じた既定(文字列は左、数値は右)に従う。</summary>
        General = 0,
        Left,
        Center,
        Right,
        Fill,
        Justify,
        CenterContinuous,
        Distributed,
    }

    /// <summary>垂直方向の配置。</summary>
    public enum VerticalAlignment
    {
        Top = 0,
        Center,
        Bottom,
        Justify,
        Distributed,
    }

    /// <summary>罫線の線種。</summary>
    public enum BorderLineStyle
    {
        None = 0,
        Thin,
        Medium,
        Thick,
        Hair,
        Dotted,
        Dashed,
        DashDot,
        DashDotDot,
        Double,
        MediumDashed,
        MediumDashDot,
        MediumDashDotDot,
        SlantDashDot,
    }

    /// <summary>下線の種類。</summary>
    public enum UnderlineStyle
    {
        None = 0,
        Single,
        Double,
        SingleAccounting,
        DoubleAccounting,
    }

    /// <summary>セルのフォント設定。</summary>
    public sealed record FontStyle(
        string Name,
        double SizePt,
        bool Bold,
        bool Italic,
        UnderlineStyle Underline,
        bool Strike,
        ArgbColor Color)
    {
        /// <summary>Excel の既定に近いフォント(テストおよびフォールバック用)。</summary>
        public static FontStyle Default { get; } =
            new("Calibri", 11.0, false, false, UnderlineStyle.None, false, ArgbColor.Black);
    }

    /// <summary>1辺分の罫線。</summary>
    public sealed record BorderEdge(BorderLineStyle Style, ArgbColor Color)
    {
        /// <summary>罫線なし。</summary>
        public static BorderEdge None { get; } = new(BorderLineStyle.None, ArgbColor.Black);

        public bool IsVisible => Style != BorderLineStyle.None;
    }

    /// <summary>セル4辺および斜めの罫線。</summary>
    public sealed record BorderSet(
        BorderEdge Left,
        BorderEdge Right,
        BorderEdge Top,
        BorderEdge Bottom,
        BorderEdge DiagonalDown,
        BorderEdge DiagonalUp)
    {
        /// <summary>すべて罫線なし。</summary>
        public static BorderSet None { get; } = new(
            BorderEdge.None, BorderEdge.None, BorderEdge.None, BorderEdge.None, BorderEdge.None, BorderEdge.None);

        public bool HasAnyVisibleEdge =>
            Left.IsVisible || Right.IsVisible || Top.IsVisible || Bottom.IsVisible
            || DiagonalDown.IsVisible || DiagonalUp.IsVisible;
    }

    /// <summary>
    /// セルの書式。要件1.2 が求める項目を保持する。
    /// </summary>
    public sealed record CellStyle(
        FontStyle Font,
        BorderSet Borders,
        HorizontalAlignment HAlign,
        VerticalAlignment VAlign,
        string? NumberFormat,
        ArgbColor BackgroundColor,
        bool WrapText,
        bool ShrinkToFit,
        int Indent)
    {
        /// <summary>既定の書式(罫線・背景なし、既定フォント)。</summary>
        public static CellStyle Default { get; } = new(
            FontStyle.Default,
            BorderSet.None,
            HorizontalAlignment.General,
            VerticalAlignment.Bottom,
            NumberFormat: null,
            BackgroundColor: ArgbColor.Transparent,
            WrapText: false,
            ShrinkToFit: false,
            Indent: 0);
    }
}
