using System;
using SkiaSharp;

namespace Utsushi.Rendering.Fonts
{
    /// <summary>
    /// ライブラリに同梱した日本語フォント(要件11.1)。
    /// </summary>
    /// <remarks>
    /// BIZ UDPゴシック(SIL Open Font License 1.1)を埋め込みリソースとして持つ。実行環境にインストールされた
    /// 同名のフォントがあればそちらを優先し(<see cref="FontResolver"/>)、無い場合にだけこれを使う。
    /// 書体はプロセス全体で1つを共有し、解放しない。
    /// </remarks>
    public static class BundledFonts
    {
        /// <summary>同梱した日本語ゴシック体のファミリ名。</summary>
        public const string JapaneseGothicFamily = "BIZ UDPGothic";

        private const string ResourceName = "Utsushi.Rendering.Fonts.BIZUDPGothic-Regular.ttf";

        // 読み込みに失敗した場合に例外をキャッシュし続けないよう、PublicationOnly にする(次の呼び出しで再試行する)。
        private static readonly Lazy<SKTypeface> JapaneseGothic = new(LoadJapaneseGothic, System.Threading.LazyThreadSafetyMode.PublicationOnly);

        /// <summary>同梱フォントのうち、指定のファミリ名(日本語名を含む)に一致するものを返す。</summary>
        internal static SKTypeface? Find(string familyName)
        {
            foreach (var candidate in FontFamilyAliases.Candidates(familyName))
            {
                if (string.Equals(candidate, JapaneseGothicFamily, StringComparison.OrdinalIgnoreCase))
                {
                    return JapaneseGothic.Value;
                }
            }

            return null;
        }

        /// <summary>同梱の日本語ゴシック体。</summary>
        internal static SKTypeface JapaneseGothicTypeface => JapaneseGothic.Value;

        private static SKTypeface LoadJapaneseGothic()
        {
            using var stream = typeof(BundledFonts).Assembly.GetManifestResourceStream(ResourceName)
                ?? throw new InvalidOperationException($"同梱フォントのリソース '{ResourceName}' が見つかりません。");
            using var data = SKData.Create(stream);
            return SKTypeface.FromData(data)
                ?? throw new InvalidOperationException($"同梱フォント '{ResourceName}' を読み込めません。");
        }
    }
}
