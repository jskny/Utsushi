using System;

namespace Utsushi.Core;

/// <summary>ポイント単位の座標。原点はページ左上、X軸は右方向、Y軸は下方向。</summary>
public readonly struct PointPt : IEquatable<PointPt>
{
    public PointPt(double x, double y)
    {
        X = x;
        Y = y;
    }

    public double X { get; }

    public double Y { get; }

    public override string ToString() => $"({X:0.###},{Y:0.###})";

    public bool Equals(PointPt other) => X.Equals(other.X) && Y.Equals(other.Y);

    public override bool Equals(object? obj) => obj is PointPt other && Equals(other);

    public override int GetHashCode() => (X.GetHashCode() * 397) ^ Y.GetHashCode();
}

/// <summary>ポイント単位の矩形。原点はページ左上、Y軸は下方向。</summary>
public readonly struct RectPt : IEquatable<RectPt>
{
    public RectPt(double left, double top, double width, double height)
    {
        if (width < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width), width, "幅は0以上である必要があります。");
        }

        if (height < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(height), height, "高さは0以上である必要があります。");
        }

        Left = left;
        Top = top;
        Width = width;
        Height = height;
    }

    public double Left { get; }

    public double Top { get; }

    public double Width { get; }

    public double Height { get; }

    public double Right => Left + Width;

    public double Bottom => Top + Height;

    public bool IsEmpty => Width <= 0 || Height <= 0;

    public static RectPt FromBounds(double left, double top, double right, double bottom) =>
        new(left, top, Math.Max(0, right - left), Math.Max(0, bottom - top));

    public RectPt Deflate(double horizontal, double vertical) =>
        FromBounds(Left + horizontal, Top + vertical, Right - horizontal, Bottom - vertical);

    public override string ToString() => $"[{Left:0.###},{Top:0.###} {Width:0.###}x{Height:0.###}]";

    public bool Equals(RectPt other) =>
        Left.Equals(other.Left) && Top.Equals(other.Top) && Width.Equals(other.Width) && Height.Equals(other.Height);

    public override bool Equals(object? obj) => obj is RectPt other && Equals(other);

    public override int GetHashCode() =>
        (((Left.GetHashCode() * 397) ^ Top.GetHashCode()) * 397 ^ Width.GetHashCode()) * 397 ^ Height.GetHashCode();
}
