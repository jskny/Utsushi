using System;
using System.Globalization;
using System.Text;

namespace Utsushi.Core
{
    /// <summary>
    /// 診断の結果(未対応要素・書式の近似)に入れる、ファイルから読んだ文字列の正規化
    /// (excel2pdf-consolidation design.md 2.1・9)。
    /// </summary>
    /// <remarks>
    /// content-type・プリセット名・表示形式のコードなどは、ファイル由来の任意の文字列になりうる。
    /// 処理ログに流れる値の大きさを抑え、改行などでログの行を崩さないよう、制御文字・書式文字を取り除いて切り詰める。
    /// </remarks>
    public static class DiagnosticText
    {
        /// <summary>正規化した文字列の最大の長さ(UTF-16 のコード単位)。</summary>
        public const int MaxLength = 64;

        /// <summary>
        /// 制御文字・書式文字(Unicode の Cc・Cf)と対になっていないサロゲートを取り除き、<see cref="MaxLength"/> で切り詰める。
        /// 切り詰めた場合は末尾に「…」を付ける。<c>null</c> は <c>null</c> のまま返す。
        /// </summary>
        public static string? Normalize(string? value)
        {
            if (value is null)
            {
                return null;
            }

            var builder = new StringBuilder(Math.Min(value.Length, MaxLength + 1));
            var truncated = false;
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (char.IsHighSurrogate(c))
                {
                    if (i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                    {
                        if (builder.Length + 2 > MaxLength)
                        {
                            truncated = true;
                            break;
                        }

                        builder.Append(c).Append(value[i + 1]);
                        i++;
                    }

                    continue;
                }

                if (char.IsLowSurrogate(c))
                {
                    continue;
                }

                var category = CharUnicodeInfo.GetUnicodeCategory(c);
                if (category == UnicodeCategory.Control || category == UnicodeCategory.Format)
                {
                    continue;
                }

                if (builder.Length + 1 > MaxLength)
                {
                    truncated = true;
                    break;
                }

                builder.Append(c);
            }

            return truncated ? builder.Append('…').ToString() : builder.ToString();
        }
    }
}
