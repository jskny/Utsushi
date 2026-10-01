using System.Collections.Generic;
using Utsushi.Parsing.Model;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// OOXML の <c>pageSetup/@paperSize</c> コードと用紙実寸の対応表。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 自社帳票で使用しうる用紙(A/B判・Letter/Legal と、はがき・封筒などの日本の用紙)のみを収録する。
    /// 未知のコードは A4 として扱う(帳票追加時に必要になったらここへ追加する)。
    /// </para>
    /// <para>
    /// コードの意味は Windows の <c>DEVMODE.dmPaperSize</c>(wingdi.h の <c>DMPAPER_*</c>)に合わせる。Excel はプリンタの
    /// 用紙コードをそのまま保存するためである。ECMA-376 Part 1(18.3.1.63 pageSetup の paperSize)の説明は一部が
    /// wingdi.h と食い違う(12/13 を ISO の B4/B5 としている)が、実際に Excel が書く値に従う。
    /// 回転(Rotated)・横送り(Transverse)の用紙は、幅と高さを wingdi.h の定義どおりに持つ
    /// (印刷の向きは <c>orientation</c> が別に決める)。
    /// </para>
    /// </remarks>
    internal static class PaperSizeTable
    {
        private static readonly Dictionary<int, PaperSize> Sizes = new()
        {
            [1] = PaperSize.FromMillimeters(1, "Letter", 215.9, 279.4),
            [5] = PaperSize.FromMillimeters(5, "Legal", 215.9, 355.6),
            [8] = PaperSize.FromMillimeters(8, "A3", 297, 420),
            [9] = PaperSize.FromMillimeters(9, "A4", 210, 297),
            [10] = PaperSize.FromMillimeters(10, "A4 Small", 210, 297),
            [11] = PaperSize.FromMillimeters(11, "A5", 148, 210),
            [12] = PaperSize.FromMillimeters(12, "B4 (JIS)", 257, 364),
            [13] = PaperSize.FromMillimeters(13, "B5 (JIS)", 182, 257),
            [42] = PaperSize.FromMillimeters(42, "B4 (ISO)", 250, 353),
            [43] = PaperSize.FromMillimeters(43, "Japanese Postcard", 100, 148),
            [55] = PaperSize.FromMillimeters(55, "A4 Transverse", 210, 297),
            [61] = PaperSize.FromMillimeters(61, "A5 Transverse", 148, 210),
            [62] = PaperSize.FromMillimeters(62, "B5 (JIS) Transverse", 182, 257),
            [66] = PaperSize.FromMillimeters(66, "A2", 420, 594),
            [67] = PaperSize.FromMillimeters(67, "A3 Transverse", 297, 420),
            [69] = PaperSize.FromMillimeters(69, "Japanese Double Postcard", 200, 148),
            [70] = PaperSize.FromMillimeters(70, "A6", 105, 148),
            [71] = PaperSize.FromMillimeters(71, "Japanese Envelope Kaku #2", 240, 332),
            [72] = PaperSize.FromMillimeters(72, "Japanese Envelope Kaku #3", 216, 277),
            [73] = PaperSize.FromMillimeters(73, "Japanese Envelope Chou #3", 120, 235),
            [74] = PaperSize.FromMillimeters(74, "Japanese Envelope Chou #4", 90, 205),
            [75] = PaperSize.FromMillimeters(75, "Letter Rotated", 279.4, 215.9),
            [76] = PaperSize.FromMillimeters(76, "A3 Rotated", 420, 297),
            [77] = PaperSize.FromMillimeters(77, "A4 Rotated", 297, 210),
            [78] = PaperSize.FromMillimeters(78, "A5 Rotated", 210, 148),
            [79] = PaperSize.FromMillimeters(79, "B4 (JIS) Rotated", 364, 257),
            [80] = PaperSize.FromMillimeters(80, "B5 (JIS) Rotated", 257, 182),
            [81] = PaperSize.FromMillimeters(81, "Japanese Postcard Rotated", 148, 100),
            [82] = PaperSize.FromMillimeters(82, "Japanese Double Postcard Rotated", 148, 200),
            [83] = PaperSize.FromMillimeters(83, "A6 Rotated", 148, 105),
            [84] = PaperSize.FromMillimeters(84, "Japanese Envelope Kaku #2 Rotated", 332, 240),
            [85] = PaperSize.FromMillimeters(85, "Japanese Envelope Kaku #3 Rotated", 277, 216),
            [86] = PaperSize.FromMillimeters(86, "Japanese Envelope Chou #3 Rotated", 235, 120),
            [87] = PaperSize.FromMillimeters(87, "Japanese Envelope Chou #4 Rotated", 205, 90),
            [88] = PaperSize.FromMillimeters(88, "B6 (JIS)", 128, 182),
            [89] = PaperSize.FromMillimeters(89, "B6 (JIS) Rotated", 182, 128),
            [91] = PaperSize.FromMillimeters(91, "Japanese Envelope You #4", 105, 235),
            [92] = PaperSize.FromMillimeters(92, "Japanese Envelope You #4 Rotated", 235, 105),
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
