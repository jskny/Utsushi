using System;
using Utsushi.Core;
using Utsushi.Layout.Model;

namespace Utsushi.Layout
{
    /// <summary>
    /// 同じページに同じ罫線を二重に描かないための、罫線の同一性の判定キー(<see cref="PageCommandBuilder"/>)。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 座標と線幅は小数点以下3桁(0.001pt)に丸めて比べる。隣接セルが同じ境界に持つ罫線は、
    /// 計算の経路の違いで座標の末尾の桁だけがずれることがあるため。
    /// </para>
    /// <para>
    /// 丸め方は、以前のキー(書式 <c>"0.###"</c> で文字列化したもの)と同じ結果になるようにしている。
    /// すなわち、有効数字15桁に丸めた値(<see cref="decimal"/> への変換と同じ)を、小数点以下3桁で
    /// 0から遠い方へ丸める。0に丸まる負の値(と -0)は、書式化すると <c>"-0"</c> になり 0 とは別のキーだったため、
    /// ここでも別の値として扱う。
    /// </para>
    /// </remarks>
    internal readonly struct BorderLineKey : IEquatable<BorderLineKey>
    {
        /// <summary>0に丸まった負の値を表す。丸めた値がこの値になることはない。</summary>
        private const long NegativeZero = long.MinValue;

        /// <summary>
        /// この大きさ以上の値は丸めずにビット列で比べる(小数点以下3桁の丸めが <see cref="long"/> に収まらないため)。
        /// 用紙上の座標・線幅がこの大きさになることはない。
        /// </summary>
        private const double MaxQuantizableMagnitude = 9e15;

        private readonly long _fromX;
        private readonly long _fromY;
        private readonly long _toX;
        private readonly long _toY;
        private readonly long _width;
        private readonly ArgbColor _color;
        private readonly LineDashStyle _dash;

        /// <summary>丸めずにビット列で比べた成分(ビットごとに fromX, fromY, toX, toY, 線幅)。</summary>
        private readonly byte _unquantized;

        public BorderLineKey(PointPt from, PointPt to, ArgbColor color, double widthPt, LineDashStyle dash)
        {
            byte unquantized = 0;
            _fromX = Quantize(from.X, 0, ref unquantized);
            _fromY = Quantize(from.Y, 1, ref unquantized);
            _toX = Quantize(to.X, 2, ref unquantized);
            _toY = Quantize(to.Y, 3, ref unquantized);
            _width = Quantize(widthPt, 4, ref unquantized);
            _unquantized = unquantized;
            _color = color;
            _dash = dash;
        }

        /// <summary>値を0.001単位の整数に丸める。</summary>
        internal static long Quantize(double value, int component, ref byte unquantized)
        {
            if (double.IsFinite(value) && Math.Abs(value) < MaxQuantizableMagnitude)
            {
                var quantized = (long)(Math.Round((decimal)value, 3, MidpointRounding.AwayFromZero) * 1000m);
                return quantized == 0 && double.IsNegative(value) ? NegativeZero : quantized;
            }

            unquantized |= (byte)(1 << component);
            return BitConverter.DoubleToInt64Bits(value);
        }

        public bool Equals(BorderLineKey other) =>
            _fromX == other._fromX
            && _fromY == other._fromY
            && _toX == other._toX
            && _toY == other._toY
            && _width == other._width
            && _unquantized == other._unquantized
            && _color.Equals(other._color)
            && _dash == other._dash;

        public override bool Equals(object? obj) => obj is BorderLineKey other && Equals(other);

        public override int GetHashCode()
        {
            var hash = default(HashCode);
            hash.Add(_fromX);
            hash.Add(_fromY);
            hash.Add(_toX);
            hash.Add(_toY);
            hash.Add(_width);
            hash.Add(_unquantized);
            hash.Add(_color);
            hash.Add(_dash);
            return hash.ToHashCode();
        }
    }
}
