using System;
using System.Globalization;
using System.Text;

namespace Utsushi.Core
{
    /// <summary>
    /// A1形式のセル番地。行・列ともに1始まり。
    /// </summary>
    public readonly struct CellAddress : IEquatable<CellAddress>, IComparable<CellAddress>
    {
        /// <summary>扱える最大列番号(Excel の XFD 列)。</summary>
        public const int MaxColumn = 16384;

        /// <summary>扱える最大行番号。</summary>
        public const int MaxRow = 1048576;

        public CellAddress(int row, int column)
        {
            if (row < 1 || row > MaxRow)
            {
                throw new ArgumentOutOfRangeException(nameof(row), row, "行番号は1以上 " + MaxRow + " 以下である必要があります。");
            }

            if (column < 1 || column > MaxColumn)
            {
                throw new ArgumentOutOfRangeException(nameof(column), column, "列番号は1以上 " + MaxColumn + " 以下である必要があります。");
            }

            Row = row;
            Column = column;
        }

        /// <summary>1始まりの行番号。</summary>
        public int Row { get; }

        /// <summary>1始まりの列番号(A=1)。</summary>
        public int Column { get; }

        /// <summary>"C3" のようなA1形式の文字列を解釈する。絶対参照記号($)は無視する。</summary>
        public static CellAddress Parse(string text)
        {
            if (!TryParse(text, out var address))
            {
                throw new FormatException($"セル番地として解釈できません: '{text}'");
            }

            return address;
        }

        /// <summary>"C3" のようなA1形式の文字列の解釈を試みる。</summary>
        public static bool TryParse(string? text, out CellAddress address)
        {
            address = default;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var s = text!.Trim().Replace("$", string.Empty);
            var i = 0;
            while (i < s.Length && IsAsciiLetter(s[i]))
            {
                i++;
            }

            if (i == 0 || i == s.Length || !TryParseColumnName(s.Substring(0, i), out var column))
            {
                return false;
            }

            if (!int.TryParse(s.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out var row))
            {
                return false;
            }

            if (row < 1 || row > MaxRow)
            {
                return false;
            }

            address = new CellAddress(row, column);
            return true;
        }

        /// <summary>"A", "Z", "AA" 形式の列名を1始まりの列番号に変換する。</summary>
        public static bool TryParseColumnName(string text, out int column)
        {
            column = 0;
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            foreach (var c in text)
            {
                if (!IsAsciiLetter(c))
                {
                    column = 0;
                    return false;
                }

                column = (column * 26) + (char.ToUpperInvariant(c) - 'A' + 1);
                if (column > MaxColumn)
                {
                    return false;
                }
            }

            return column >= 1;
        }

        /// <summary>
        /// 列名に使える英字('A'..'Z'・'a'..'z')かどうか。
        /// </summary>
        /// <remarks>
        /// <see cref="char.IsLetter(char)"/> はASCII以外の英字(é・全角Ａなど)も受け付けるため、
        /// 列番号の計算(<c>c - 'A' + 1</c>)が範囲外の値になり、"é1" が EG1 のような無関係なセルとして解釈されてしまう。
        /// </remarks>
        private static bool IsAsciiLetter(char c) => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');

        /// <summary>1始まりの列番号を "A", "Z", "AA" 形式の列名に変換する。</summary>
        public static string ColumnName(int column)
        {
            if (column < 1 || column > MaxColumn)
            {
                throw new ArgumentOutOfRangeException(nameof(column), column, "列番号が範囲外です。");
            }

            var sb = new StringBuilder(3);
            var n = column;
            while (n > 0)
            {
                var rem = (n - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                n = (n - 1) / 26;
            }

            return sb.ToString();
        }

        public override string ToString() => ColumnName(Column) + Row.ToString(CultureInfo.InvariantCulture);

        public bool Equals(CellAddress other) => Row == other.Row && Column == other.Column;

        public override bool Equals(object? obj) => obj is CellAddress other && Equals(other);

        public override int GetHashCode() => (Row * 397) ^ Column;

        /// <summary>行優先(上から下、同一行では左から右)で比較する。</summary>
        public int CompareTo(CellAddress other)
        {
            var byRow = Row.CompareTo(other.Row);
            return byRow != 0 ? byRow : Column.CompareTo(other.Column);
        }

        public static bool operator ==(CellAddress left, CellAddress right) => left.Equals(right);

        public static bool operator !=(CellAddress left, CellAddress right) => !left.Equals(right);
    }
}
