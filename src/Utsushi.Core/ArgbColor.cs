using System;
using System.Globalization;

namespace Utsushi.Core
{
    /// <summary>
    /// 8bit ARGB の色。レイヤー間で色を受け渡すための共通表現(SkiaSharp等の描画ライブラリ型に依存しない)。
    /// </summary>
    public readonly struct ArgbColor : IEquatable<ArgbColor>
    {
        public ArgbColor(byte a, byte r, byte g, byte b)
        {
            A = a;
            R = r;
            G = g;
            B = b;
        }

        public byte A { get; }

        public byte R { get; }

        public byte G { get; }

        public byte B { get; }

        public static ArgbColor Black => new(0xFF, 0x00, 0x00, 0x00);

        public static ArgbColor White => new(0xFF, 0xFF, 0xFF, 0xFF);

        /// <summary>完全に透明な色。塗りつぶし不要を表す。</summary>
        public static ArgbColor Transparent => new(0x00, 0x00, 0x00, 0x00);

        public bool IsTransparent => A == 0;

        /// <summary>RGB はそのままで、アルファだけを <paramref name="alpha"/> にした色を返す。</summary>
        public ArgbColor WithAlpha(byte alpha) => new(alpha, R, G, B);

        /// <summary>"FF0000" / "FFFF0000" / "#FF0000" 形式の16進表記を解釈する。アルファ省略時は不透明として扱う。</summary>
        public static bool TryParseHex(string? text, out ArgbColor color)
        {
            color = default;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var s = text!.Trim();
            if (s.StartsWith("#", StringComparison.Ordinal))
            {
                s = s.Substring(1);
            }

            if (s.Length != 6 && s.Length != 8)
            {
                return false;
            }

            if (!uint.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value))
            {
                return false;
            }

            if (s.Length == 6)
            {
                value |= 0xFF000000;
            }

            color = new ArgbColor(
                (byte)((value >> 24) & 0xFF),
                (byte)((value >> 16) & 0xFF),
                (byte)((value >> 8) & 0xFF),
                (byte)(value & 0xFF));
            return true;
        }

        public override string ToString() =>
            "#" + A.ToString("X2", CultureInfo.InvariantCulture)
                + R.ToString("X2", CultureInfo.InvariantCulture)
                + G.ToString("X2", CultureInfo.InvariantCulture)
                + B.ToString("X2", CultureInfo.InvariantCulture);

        public bool Equals(ArgbColor other) => A == other.A && R == other.R && G == other.G && B == other.B;

        public override bool Equals(object? obj) => obj is ArgbColor other && Equals(other);

        public override int GetHashCode() => (((A * 397) ^ R) * 397 ^ G) * 397 ^ B;

        public static bool operator ==(ArgbColor left, ArgbColor right) => left.Equals(right);

        public static bool operator !=(ArgbColor left, ArgbColor right) => !left.Equals(right);
    }
}
