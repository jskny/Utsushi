using System;

namespace Utsushi.Core
{
    /// <summary>
    /// 矩形のセル範囲(両端を含む)。
    /// </summary>
    public readonly struct CellRange : IEquatable<CellRange>
    {
        public CellRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            if (firstRow < 1 || firstColumn < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(firstRow), "行・列番号は1始まりです。");
            }

            if (lastRow < firstRow || lastColumn < firstColumn)
            {
                throw new ArgumentException("範囲の終端は始端以上である必要があります。", nameof(lastRow));
            }

            FirstRow = firstRow;
            FirstColumn = firstColumn;
            LastRow = lastRow;
            LastColumn = lastColumn;
        }

        public CellRange(CellAddress first, CellAddress last)
            : this(
                Math.Min(first.Row, last.Row),
                Math.Min(first.Column, last.Column),
                Math.Max(first.Row, last.Row),
                Math.Max(first.Column, last.Column))
        {
        }

        public int FirstRow { get; }

        public int FirstColumn { get; }

        public int LastRow { get; }

        public int LastColumn { get; }

        public int RowCount => LastRow - FirstRow + 1;

        public int ColumnCount => LastColumn - FirstColumn + 1;

        public CellAddress TopLeft => new(FirstRow, FirstColumn);

        public CellAddress BottomRight => new(LastRow, LastColumn);

        /// <summary>"A1:D10" 形式(単一セル "A1" も可)の範囲文字列を解釈する。</summary>
        public static CellRange Parse(string text)
        {
            if (!TryParse(text, out var range))
            {
                throw new FormatException($"セル範囲として解釈できません: '{text}'");
            }

            return range;
        }

        /// <summary>"A1:D10" 形式(単一セル "A1" も可)の範囲文字列の解釈を試みる。</summary>
        public static bool TryParse(string? text, out CellRange range)
        {
            range = default;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var parts = text!.Split(':');
            if (parts.Length == 1)
            {
                if (!CellAddress.TryParse(parts[0], out var single))
                {
                    return false;
                }

                range = new CellRange(single, single);
                return true;
            }

            if (parts.Length != 2)
            {
                return false;
            }

            if (!CellAddress.TryParse(parts[0], out var first) || !CellAddress.TryParse(parts[1], out var last))
            {
                return false;
            }

            range = new CellRange(first, last);
            return true;
        }

        public bool Contains(CellAddress address) =>
            address.Row >= FirstRow && address.Row <= LastRow &&
            address.Column >= FirstColumn && address.Column <= LastColumn;

        public bool Contains(int row, int column) =>
            row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;

        public override string ToString() => TopLeft + ":" + BottomRight;

        public bool Equals(CellRange other) =>
            FirstRow == other.FirstRow && FirstColumn == other.FirstColumn &&
            LastRow == other.LastRow && LastColumn == other.LastColumn;

        public override bool Equals(object? obj) => obj is CellRange other && Equals(other);

        public override int GetHashCode() => (((FirstRow * 397) ^ FirstColumn) * 397 ^ LastRow) * 397 ^ LastColumn;

        public static bool operator ==(CellRange left, CellRange right) => left.Equals(right);

        public static bool operator !=(CellRange left, CellRange right) => !left.Equals(right);
    }
}
