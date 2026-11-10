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
        public static string? Normalize(string? value) => Normalize(value, MaxLength);

        /// <summary>
        /// <see cref="Normalize(string)"/> と同じく制御文字・書式文字と孤立したサロゲートを取り除き、<paramref name="maxLength"/> で切り詰める。
        /// </summary>
        public static string? Normalize(string? value, int maxLength)
        {
            if (value is null)
            {
                return null;
            }

            var builder = new StringBuilder(Math.Min(value.Length, Math.Min(maxLength, 4096) + 1));
            var truncated = false;
            for (var i = 0; i < value.Length; i++)
            {
                var c = value[i];
                if (char.IsHighSurrogate(c))
                {
                    if (i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                    {
                        if (builder.Length + 2 > maxLength)
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
                if (IsRemoved(category))
                {
                    continue;
                }

                if (builder.Length + 1 > maxLength)
                {
                    truncated = true;
                    break;
                }

                builder.Append(c);
            }

            return truncated ? builder.Append('…').ToString() : builder.ToString();
        }

        /// <summary>
        /// 診断の文から取り除く文字の種類。制御文字・書式文字に加え、行区切り・段落区切り(U+2028・U+2029)も取り除く
        /// (改行として扱うログのビューアで、偽の行を差し込ませないため)。
        /// </summary>
        public static bool IsRemoved(UnicodeCategory category) =>
            category is UnicodeCategory.Control or UnicodeCategory.Format
                or UnicodeCategory.LineSeparator or UnicodeCategory.ParagraphSeparator;

        /// <summary>Excel のシート名の最大の長さ。</summary>
        public const int MaxSheetNameLength = 31;

        /// <summary>文書名の最大の長さ。</summary>
        public const int MaxDocumentNameLength = 255;

        private static readonly char[] InvalidSheetNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

        /// <summary>
        /// シート名が Excel の制約(空でない、31文字以下、<c>: \ / ? * [ ]</c> と制御文字・書式文字・行区切りを含まない)を満たすかどうか。
        /// </summary>
        /// <remarks>
        /// シート名はエラーのメッセージ・診断の結果に入り、処理ログに流れる。ファイルから読んだ名前と、呼び出し元が指定した名前の
        /// 両方をこの規則で確かめる(security-reviewer指摘)。
        /// </remarks>
        public static bool IsValidSheetName(string? name) =>
            !string.IsNullOrEmpty(name)
            && name!.Length <= MaxSheetNameLength
            && name.IndexOfAny(InvalidSheetNameChars) < 0
            && !ContainsRemoved(name);

        /// <summary>
        /// 文書名として使えるかどうか(255文字以下、制御文字・書式文字・行区切りを含まない)。文書名は PDF のタイトル・
        /// ヘッダー/フッターの <c>&amp;F</c>・エラー情報に入る。
        /// </summary>
        public static bool IsValidDocumentName(string? name) =>
            name is not null && name.Length <= MaxDocumentNameLength && !ContainsRemoved(name);

        private static bool ContainsRemoved(string text)
        {
            foreach (var c in text)
            {
                if (IsRemoved(CharUnicodeInfo.GetUnicodeCategory(c)))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
