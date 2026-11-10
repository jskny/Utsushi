using System;

namespace Utsushi.Parsing.OpenXml
{
    /// <summary>
    /// 診断(書式の近似)にしか使わない属性を読む。OpenXml SDK の型付き属性(<c>.Value</c>)は不正な文字列で
    /// <see cref="FormatException"/>(範囲外の数値は <see cref="OverflowException"/>)を投げるが、表示に使わない値のためにファイル全体を破損として扱わないよう、既定値にする
    /// (要件6.12と同じく、不正な値は既定値。excel2pdf-consolidation 要件7.2)。
    /// </summary>
    internal static class LenientValue
    {
        public static T Read<T>(Func<T> read, T fallback)
        {
            try
            {
                return read();
            }
            catch (Exception ex) when (ex is FormatException or OverflowException)
            {
                return fallback;
            }
        }
    }
}
