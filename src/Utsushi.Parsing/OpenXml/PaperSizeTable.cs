using System.Collections.Generic;
using Utsushi.Parsing.Model;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// OOXML の <c>pageSetup/@paperSize</c> コードと用紙実寸の対応表。
    /// </summary>
    /// <remarks>
    /// 自社帳票で使用しうる用紙(A/B判および Letter/Legal)のみを収録する。
    /// 未知のコードは A4 として扱う(帳票追加時に必要になったらここへ追加する)。
    /// </remarks>
    internal static class PaperSizeTable
    {
        private static readonly Dictionary<int, PaperSize> Sizes = new()
        {
            [1] = PaperSize.FromMillimeters(1, "Letter", 215.9, 279.4),
            [5] = PaperSize.FromMillimeters(5, "Legal", 215.9, 355.6),
            [8] = PaperSize.FromMillimeters(8, "A3", 297, 420),
            [9] = PaperSize.FromMillimeters(9, "A4", 210, 297),
            [11] = PaperSize.FromMillimeters(11, "A5", 148, 210),
            [12] = PaperSize.FromMillimeters(12, "B4 (JIS)", 257, 364),
            [13] = PaperSize.FromMillimeters(13, "B5 (JIS)", 182, 257),
            [43] = PaperSize.FromMillimeters(43, "B4 (JIS) 2", 257, 364),
            [62] = PaperSize.FromMillimeters(62, "B4 (JIS) Rotated", 364, 257),
            [66] = PaperSize.FromMillimeters(66, "A2", 420, 594),
            [70] = PaperSize.FromMillimeters(70, "A6", 105, 148),
        };

        /// <summary>用紙コードから用紙サイズを引く。未知のコードは A4 を返す。</summary>
        public static PaperSize Resolve(int? code)
        {
            if (code is { } value && Sizes.TryGetValue(value, out var size))
            {
                return size;
            }

            return PaperSize.A4;
        }
    }
}
