using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SkiaSharp;
using Utsushi.Core.Exceptions;
using Utsushi.Rendering.Fonts;
using Utsushi.Parsing.Model;

namespace Utsushi.Rendering
{
    /// <summary>
    /// 帳票が使用するフォント名を <see cref="SKTypeface"/> へ解決する。
    /// </summary>
    /// <remarks>
    /// <para>
    /// 実行時に別フォントへ暗黙フォールバックすると帳票の見た目が崩れるため、
    /// 既定(<see cref="FontResolverOptions.Strict"/>)では解決できないフォントを例外にする
    /// (design.md「Rendering レイヤー」)。
    /// ただし、フォント自体は解決できたうえで、その書体に字形が無い個々の文字(外字)は、厳格モードでも
    /// <see cref="FontResolverOptions.GlyphFallbackFamilies"/> の書体で補う(要件11.2)。これを行わない場合は一覧を空にする。
    /// </para>
    /// <para>
    /// 明示的に登録したフォントファイル(<c>.ttf</c>/<c>.otf</c>/<c>.ttc</c>)を最優先で使い、
    /// 次に実行環境にインストールされたフォントを探す。
    /// </para>
    /// <para>
    /// 太字・斜体は、実フォントがその字形を持つ場合のみ実字形を使う。持たない場合は
    /// 通常字形の書体を返し、<see cref="ResolvedTypeface.SynthesizeBold"/> /
    /// <see cref="ResolvedTypeface.SynthesizeItalic"/> を立てて描画側に装飾を委ねる。
    /// SkiaSharp に合成させると PDF が Type 3 フォントになり、文字列検索ができなくなるため。
    /// </para>
    /// </remarks>
    public sealed class FontResolver : IDisposable
    {
        /// <summary>フォントの同一性判定に使う先頭バイト数。</summary>
        private const int FontIdentityProbeBytes = 4096;

        private readonly FontResolverOptions _options;
        private readonly Dictionary<FontKey, SKTypeface> _registered = new();
        private readonly ConcurrentDictionary<FontKey, ResolvedTypeface> _cache = new();
        private readonly ConcurrentDictionary<(bool Bold, bool Italic), IReadOnlyList<ResolvedTypeface>> _glyphFallbacks = new();
        private readonly List<SKTypeface> _owned = new();
        private readonly object _registrationLock = new();
        private readonly object _resolveLock = new();
        private readonly HashSet<FontKey> _substituted = new();
        private bool _disposed;

        public FontResolver(FontResolverOptions? options = null)
        {
            _options = options ?? FontResolverOptions.Strict;

            foreach (var (key, path) in _options.FontFiles)
            {
                RegisterFontFile(key, path);
            }

            foreach (var directory in _options.FontDirectories)
            {
                RegisterFontDirectory(directory);
            }
        }

        /// <summary>
        /// フォント名・太字・斜体から書体を解決する。
        /// </summary>
        /// <exception cref="FontNotAvailableException">
        /// 厳格モードでフォントが見つからない場合。
        /// </exception>
        public ResolvedTypeface Resolve(FontStyle font)
        {
            ThrowIfDisposed();

            // 複数スレッドから同時に変換した場合でも、同じ書体を二重に生成しないよう、生成はロックの内側で1回だけ行う。
            return Resolve(font, optional: false)!;
        }

        /// <summary>
        /// セルのフォントに字形が無い文字を描画するための、外字用の代替書体を優先順に返す(要件11.2)。
        /// </summary>
        /// <remarks>
        /// <see cref="FontResolverOptions.GlyphFallbackFamilies"/> の各フォントを、明示登録 → インストール済み →
        /// 同梱の順に探す。見つからないフォントは飛ばす(IPAmj明朝のように、入っていれば使うフォントを既定に含めるため)。
        /// 太字・斜体は、代替書体に実字形が無ければセルのフォントと同様に描画側で合成する。
        /// </remarks>
        public IReadOnlyList<ResolvedTypeface> ResolveGlyphFallbacks(bool bold, bool italic)
        {
            ThrowIfDisposed();
            return _glyphFallbacks.GetOrAdd((bold, italic), key =>
            {
                var list = new List<ResolvedTypeface>();
                foreach (var family in _options.GlyphFallbackFamilies)
                {
                    var resolved = Resolve(FontStyle.Default with { Name = family, Bold = key.Bold, Italic = key.Italic }, optional: true);
                    if (resolved is not null && !list.Exists(r => ReferenceEquals(r.Typeface, resolved.Typeface)))
                    {
                        list.Add(resolved);
                    }
                }

                return list;
            });
        }

        private ResolvedTypeface? Resolve(FontStyle font, bool optional)
        {
            var key = new FontKey(font.Name, font.Bold, font.Italic);
            if (_cache.TryGetValue(key, out var cached))
            {
                // 見つからずに代替フォントで解決した結果は、外字用の代替フォントとしては使わない
                // (「IPAmjMincho」の名前で別の書体を拾い、処理の順序で結果が変わるのを防ぐ)。
                return optional && IsSubstituted(key) ? null : cached;
            }

            lock (_resolveLock)
            {
                ThrowIfDisposed();
                if (_cache.TryGetValue(key, out cached))
                {
                    return optional && _substituted.Contains(key) ? null : cached;
                }

                var resolved = ResolveCore(font);
                if (resolved is null)
                {
                    if (optional)
                    {
                        return null;
                    }

                    resolved = HandleMissingFont(font);
                    _substituted.Add(key);
                }

                _cache[key] = resolved;
                return resolved;
            }
        }

        private bool IsSubstituted(FontKey key)
        {
            lock (_resolveLock)
            {
                return _substituted.Contains(key);
            }
        }

        private ResolvedTypeface? ResolveCore(FontStyle font)
        {
            // 1. 明示登録されたフォントファイル(スタイル指定つき)を最優先で使う。
            if (TryGetRegistered(font.Name, font.Bold, font.Italic, out var registeredExact))
            {
                return new ResolvedTypeface(registeredExact, SynthesizeBold: false, SynthesizeItalic: false);
            }

            // 2. 明示登録の通常字形があれば、それを使って装飾を合成する。
            if (TryGetRegistered(font.Name, bold: false, italic: false, out var registeredRegular))
            {
                return new ResolvedTypeface(registeredRegular, font.Bold, font.Italic);
            }

            // 3. 実行環境にインストールされたフォントを探す。無ければ同梱フォント(要件11.1)を探す。
            var regular = FindInstalled(font.Name, bold: false, italic: false);
            if (regular is null)
            {
                return BundledFonts.Find(font.Name) is { } bundled
                    ? new ResolvedTypeface(bundled, font.Bold, font.Italic)
                    : null;
            }

            if (!font.Bold && !font.Italic)
            {
                return new ResolvedTypeface(regular, false, false);
            }

            // 太字/斜体が要求された場合、実字形を持つ別ファイルがあるかを確認する。
            // SKTypeface.FromFamilyName は字形が無くても合成した書体を返すため、
            // 通常字形とフォントデータが同一かどうかで「実字形の有無」を判定する。
            // インストール済みの書体はプロセス全体で共有されるマネージドオブジェクトのため、使わなかった側も解放しない
            // (解放すると、キャッシュ済みの通常字形と同じインスタンスを解放しうる)。
            var styled = FindInstalled(font.Name, font.Bold, font.Italic);
            if (styled is not null && !IsSameFontData(regular, styled))
            {
                // 別ファイルでも、要求した装飾の一部しか持たないことがある(例: Bold はあるが Bold Italic は無く、
                // fontconfig が Bold に斜体の傾きを付けて返す)。その傾きは書体のデータには含まれず、サブセット化
                // (要件11.5)で作り直した書体からは失われるため、フォントのデータ自体が持つ装飾を見て、
                // 足りない装飾だけを描画側で合成する。
                var (hasBold, hasItalic) = ReadEmbeddedStyle(styled);
                return new ResolvedTypeface(styled, font.Bold && !hasBold, font.Italic && !hasItalic);
            }

            return new ResolvedTypeface(regular, font.Bold, font.Italic);
        }

        /// <summary>
        /// フォントのデータ自体が太字・斜体の字形かどうかを、OS/2 表(無ければ head 表)の書体情報から読む。
        /// </summary>
        /// <remarks>
        /// <c>SKTypeface.FontStyle</c> は fontconfig が合成した装飾(斜体の傾きなど)も含めた要求どおりの値を返すため使わない。
        /// </remarks>
        internal static (bool Bold, bool Italic) ReadEmbeddedStyle(SKTypeface typeface)
        {
            const uint Os2Tag = 0x4F532F32;
            const uint HeadTag = 0x68656164;
            try
            {
                // SkiaSharp の GetTableData は、表を読めない場合に基底の Exception を投げるため TryGetTableData を使う。
                if (typeface.TryGetTableData(Os2Tag, out var os2) && os2 is not null)
                {
                    if (os2.Length >= 64)
                    {
                        var weightClass = (os2[4] << 8) | os2[5];
                        var fsSelection = (os2[62] << 8) | os2[63];
                        return (weightClass >= 600 || (fsSelection & 0x0020) != 0, (fsSelection & 0x0201) != 0);
                    }
                }

                if (typeface.TryGetTableData(HeadTag, out var head) && head is not null)
                {
                    if (head.Length >= 46)
                    {
                        var macStyle = (head[44] << 8) | head[45];
                        return ((macStyle & 0x1) != 0, (macStyle & 0x2) != 0);
                    }
                }
            }
            catch (Exception ex) when (ex is ArgumentException or IndexOutOfRangeException)
            {
                // 読めない場合は下の既定値(要求どおりの実字形とみなす。従来の判定と同じ)に倒す。
            }

            return (true, true);
        }

        private ResolvedTypeface HandleMissingFont(FontStyle font)
        {
            if (_options.Mode == FontResolutionMode.Strict)
            {
                throw new FontNotAvailableException(
                    font.Name,
                    $"フォント '{font.Name}' が実行環境に見つかりません。"
                    + "帳票の見た目が崩れるため代替フォントへのフォールバックは行いません。"
                    + "フォントをインストールするか、FontResolverOptions でフォントファイルを登録してください。");
            }

            if (_options.FallbackFamilyName is { } fallbackName)
            {
                var fallback = FindInstalled(fallbackName, bold: false, italic: false) ?? BundledFonts.Find(fallbackName);
                if (fallback is not null)
                {
                    return new ResolvedTypeface(fallback, font.Bold, font.Italic);
                }
            }

            // 代替フォント名の指定が無い(または見つからない)場合は、実行環境の既定書体(Linuxでは多くの場合
            // 日本語の字形を持たない DejaVu Sans)ではなく、同梱の日本語フォントを使う(要件11.1)。
            return new ResolvedTypeface(BundledFonts.JapaneseGothicTypeface, font.Bold, font.Italic);
        }

        /// <summary>
        /// インストール済みフォントから、要求したファミリ名と厳密に一致する書体を探す。
        /// </summary>
        /// <remarks>
        /// SkiaSharp は未知のファミリ名に対しても既定フォントを返すため、
        /// 得られた書体のファミリ名を照合して「本当に要求したフォントか」を確認する。
        /// </remarks>
        private static SKTypeface? FindInstalled(string familyName, bool bold, bool italic)
        {
            var style = new SKFontStyle(
                bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal,
                italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);

            var typeface = SKTypeface.FromFamilyName(familyName, style);
            if (typeface is null)
            {
                return null;
            }

            if (FontFamilyAliases.AreSame(familyName, typeface.FamilyName))
            {
                return typeface;
            }

            // 日本語名(「ＭＳ Ｐゴシック」等)で見つからない場合は英語名でも探す(要件11.6)。逆も同様。
            foreach (var alias in FontFamilyAliases.Candidates(familyName).Skip(1))
            {
                var aliased = SKTypeface.FromFamilyName(alias, style);
                if (aliased is not null && FontFamilyAliases.AreSame(alias, aliased.FamilyName))
                {
                    return aliased;
                }
            }

            return null;
        }

        /// <summary>
        /// 2つの書体が同一のフォントデータかどうかを判定する。
        /// </summary>
        /// <remarks>
        /// 同一なら、要求した太字/斜体の字形をフォントが持たず SkiaSharp が合成したことを意味する。
        /// 全バイトの比較はフォントサイズ(日本語フォントで数MB)を考えると重いため、
        /// バイト長と先頭 <see cref="FontIdentityProbeBytes"/> バイトで判定する。
        /// 先頭にはテーブルディレクトリが含まれ、別ファイルであればまず一致しない。
        /// </remarks>
        private static bool IsSameFontData(SKTypeface left, SKTypeface right)
        {
            using var leftStream = left.OpenStream();
            using var rightStream = right.OpenStream();

            if (leftStream is null || rightStream is null)
            {
                // データを読めない場合は「別物と判断できない」ため、合成扱いにして安全側へ倒す。
                return true;
            }

            if (leftStream.Length != rightStream.Length)
            {
                return false;
            }

            var leftHead = ReadHead(leftStream);
            var rightHead = ReadHead(rightStream);

            if (leftHead.Length != rightHead.Length)
            {
                return false;
            }

            for (var i = 0; i < leftHead.Length; i++)
            {
                if (leftHead[i] != rightHead[i])
                {
                    return false;
                }
            }

            return true;
        }

        private static byte[] ReadHead(SKStreamAsset stream)
        {
            var length = Math.Min(FontIdentityProbeBytes, stream.Length);
            var buffer = new byte[length];
            stream.Read(buffer, buffer.Length);
            return buffer;
        }

        /// <summary>
        /// フォントファイルを、フォントキーに紐づけて登録する。
        /// </summary>
        /// <param name="fontKey">
        /// <c>ファミリ名</c>、または <c>ファミリ名:bold</c> / <c>:italic</c> / <c>:bolditalic</c>。
        /// スタイル指定なしで登録したフォントは、太字/斜体が要求された際に描画側で合成される。
        /// </param>
        /// <param name="path">フォントファイルのパス。</param>
        public void RegisterFontFile(string fontKey, string path)
        {
            ThrowIfDisposed();

            var key = FontKey.Parse(fontKey);

            if (!File.Exists(path))
            {
                throw new FontNotAvailableException(
                    key.Name, $"登録しようとしたフォントファイルが存在しません: {path}");
            }

            var typeface = SKTypeface.FromFile(path)
                ?? throw new FontNotAvailableException(
                    key.Name, $"フォントファイルを読み込めません(対応していない形式の可能性があります): {path}");

            lock (_registrationLock)
            {
                _owned.Add(typeface);
                _registered[key] = typeface;
            }
        }

        /// <summary>ディレクトリ内のフォントファイルを、その内部のファミリ名で一括登録する。</summary>
        public void RegisterFontDirectory(string directory)
        {
            ThrowIfDisposed();

            if (!Directory.Exists(directory))
            {
                return;
            }

            foreach (var path in Directory.EnumerateFiles(directory))
            {
                var extension = Path.GetExtension(path);
                if (!extension.Equals(".ttf", StringComparison.OrdinalIgnoreCase)
                    && !extension.Equals(".otf", StringComparison.OrdinalIgnoreCase)
                    && !extension.Equals(".ttc", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var typeface = SKTypeface.FromFile(path);
                if (typeface is null)
                {
                    continue;
                }

                // ファイル自身が申告するスタイルで登録する。太字ファイルは太字キーに入る。
                var key = new FontKey(typeface.FamilyName, typeface.IsBold, typeface.IsItalic);

                lock (_registrationLock)
                {
                    _owned.Add(typeface);
                    _registered[key] = typeface;
                }
            }
        }

        /// <summary>
        /// 指定されたフォントがすべて解決できることを事前に検証する。
        /// </summary>
        /// <remarks>
        /// デプロイ時・起動時にフォント不足を検出するための入口(タスク6.2)。
        /// 実際の変換処理の前に呼び出すことで、描画途中での失敗を避けられる。
        /// </remarks>
        /// <exception cref="FontNotAvailableException">解決できないフォントがある場合。</exception>
        public void EnsureAvailable(IEnumerable<FontStyle> fonts)
        {
            foreach (var font in fonts)
            {
                Resolve(font);
            }
        }

        private bool TryGetRegistered(string familyName, bool bold, bool italic, out SKTypeface typeface)
        {
            lock (_registrationLock)
            {
                // 日本語名 ↔ 英語名の別名でも探す(要件11.6。例: 「MS PGothic」で登録した msgothic.ttc を
                // 日本語版Excelの「ＭＳ Ｐゴシック」で引く)。
                foreach (var candidate in FontFamilyAliases.Candidates(familyName))
                {
                    if (_registered.TryGetValue(new FontKey(candidate, bold, italic), out typeface!))
                    {
                        return true;
                    }
                }

                typeface = null!;
                return false;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            lock (_resolveLock)
            {
                _disposed = true;

                lock (_registrationLock)
                {
                    // インストール済みフォントから解決した書体は解放しない。SkiaSharp は同じネイティブ書体に
                    // 対してプロセス全体で同一のマネージドオブジェクトを返すため、ここで解放すると
                    // 同じプロセス内の別の FontResolver(別のコンバータ)が使っている書体まで壊してしまう。
                    foreach (var typeface in _owned)
                    {
                        typeface.Dispose();
                    }

                    _owned.Clear();
                    _registered.Clear();
                }

                _cache.Clear();
                _glyphFallbacks.Clear();
                _substituted.Clear();
            }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(FontResolver));
            }
        }

        /// <summary>ファミリ名とスタイルの組。ファミリ名の大小は区別しない。</summary>
        private readonly struct FontKey : IEquatable<FontKey>
        {
            public FontKey(string name, bool bold, bool italic)
            {
                Name = name;
                Bold = bold;
                Italic = italic;
            }

            public string Name { get; }

            public bool Bold { get; }

            public bool Italic { get; }

            /// <summary>"MS PGothic" / "MS PGothic:bold" 形式のキー文字列を解釈する。</summary>
            public static FontKey Parse(string fontKey)
            {
                if (string.IsNullOrWhiteSpace(fontKey))
                {
                    throw new ArgumentException("フォントキーが空です。", nameof(fontKey));
                }

                var separator = fontKey.LastIndexOf(':');
                if (separator <= 0)
                {
                    return new FontKey(fontKey.Trim(), false, false);
                }

                var name = fontKey.Substring(0, separator).Trim();
                var style = fontKey.Substring(separator + 1).Trim().ToLowerInvariant();

                return style switch
                {
                    "bold" => new FontKey(name, true, false),
                    "italic" => new FontKey(name, false, true),
                    "bolditalic" => new FontKey(name, true, true),
                    "regular" or "" => new FontKey(name, false, false),
                    _ => throw new ArgumentException(
                        $"フォントキーのスタイル '{style}' は不正です(bold / italic / bolditalic / regular)。",
                        nameof(fontKey)),
                };
            }

            public bool Equals(FontKey other) =>
                Bold == other.Bold
                && Italic == other.Italic
                && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase);

            public override bool Equals(object? obj) => obj is FontKey other && Equals(other);

            public override int GetHashCode() =>
                ((StringComparer.OrdinalIgnoreCase.GetHashCode(Name) * 397) ^ Bold.GetHashCode()) * 397
                ^ Italic.GetHashCode();
        }
    }

    /// <summary>フォント解決の方針。</summary>
    public enum FontResolutionMode
    {
        /// <summary>解決できないフォントを例外にする(既定)。</summary>
        Strict = 0,

        /// <summary>解決できない場合に代替フォントを使う。見た目が崩れる可能性がある。</summary>
        AllowFallback,
    }

    /// <summary>
    /// <see cref="FontResolver"/> の設定。
    /// </summary>
    /// <param name="Mode">フォント解決の方針。</param>
    /// <param name="FontFiles">
    /// フォントキー → フォントファイルパスの明示的な対応。
    /// キーは <c>ファミリ名</c> または <c>ファミリ名:bold</c> / <c>:italic</c> / <c>:bolditalic</c>。
    /// </param>
    /// <param name="FontDirectories">フォントファイルを一括登録するディレクトリ。</param>
    /// <param name="FallbackFamilyName">
    /// <see cref="FontResolutionMode.AllowFallback"/> のときに使う代替フォント名。null(または見つからない)の場合は
    /// 同梱の日本語フォント(<see cref="BundledFonts.JapaneseGothicFamily"/>)を使う(要件11.1)。
    /// </param>
    public sealed record FontResolverOptions(
        FontResolutionMode Mode,
        IReadOnlyDictionary<string, string> FontFiles,
        IReadOnlyList<string> FontDirectories,
        string? FallbackFamilyName)
    {
        /// <summary>
        /// 外字用の代替フォントの既定の一覧(要件11.2)。同梱のBIZ UDPゴシック(JIS第1〜第4水準)→
        /// IPAmj明朝(文字情報基盤のMJ文字。インストールされている場合のみ)の順。
        /// </summary>
        public static IReadOnlyList<string> DefaultGlyphFallbackFamilies { get; } =
            new[] { BundledFonts.JapaneseGothicFamily, "IPAmjMincho" };

        /// <summary>
        /// セルのフォントに字形が無い文字を描画する外字用の代替フォント名を、優先順に並べたもの(要件11.2)。
        /// </summary>
        /// <remarks>
        /// 既定は <see cref="DefaultGlyphFallbackFamilies"/>。社内の外字フォントを使う場合は、そのファイルを
        /// <see cref="FontFiles"/> に登録し、ここにそのフォントキーを加える。空にすると文字単位の代替を行わない
        /// (セルのフォントに字形が無い文字は要件5.5のエラーになる)。
        /// </remarks>
        public IReadOnlyList<string> GlyphFallbackFamilies { get; init; } = DefaultGlyphFallbackFamilies;

        /// <summary>厳格モード(既定)。</summary>
        public static FontResolverOptions Strict { get; } = new(
            FontResolutionMode.Strict,
            new Dictionary<string, string>(),
            Array.Empty<string>(),
            null);

        /// <summary>フォールバックを許容する設定。見た目の再現性より処理の継続を優先する場合に使う。</summary>
        public static FontResolverOptions AllowFallback(string? fallbackFamilyName = null) => new(
            FontResolutionMode.AllowFallback,
            new Dictionary<string, string>(),
            Array.Empty<string>(),
            fallbackFamilyName);
    }
}
