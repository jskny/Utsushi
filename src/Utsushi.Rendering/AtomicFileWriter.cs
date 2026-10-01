using System;
using System.IO;
using Utsushi.Core.Exceptions;

namespace Utsushi.Rendering
{
    /// <summary>
    /// ストリームの内容を、同一ディレクトリ上の一時ファイル経由でファイルへ書き出す。
    /// </summary>
    /// <remarks>
    /// 書き込み途中で異常終了しても、出力先には旧ファイル(または何も)が残る(要件5.4)。
    /// <see cref="IPdfRenderer"/> の実装によらず同じ保証を提供するため、<see cref="SkiaPdfRenderer"/>
    /// とファサード(<c>Utsushi.ReportPdfConverter</c>)の両方から使う共通処理として切り出している。
    /// </remarks>
    public static class AtomicFileWriter
    {
        /// <summary>
        /// <paramref name="content"/> の内容を <paramref name="path"/> へアトミックに書き出す。
        /// </summary>
        /// <param name="path">出力先のファイルパス。</param>
        /// <param name="content">書き出す内容(呼び出し側で完全に用意済みのストリーム)。</param>
        /// <param name="reportCode">エラーに含める帳票コード(判明している場合)。</param>
        /// <param name="sheetName">エラーに含めるシート名(判明している場合)。</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> または <paramref name="content"/> が null の場合。</exception>
        /// <exception cref="PdfRenderingException">
        /// 書き込みに失敗した場合。出力先のパスが不正な場合(NUL文字を含む、途中に既存のファイルがある等)や、
        /// 出力先ディレクトリを作成できない場合も含む。
        /// </exception>
        public static void Write(string path, Stream content, string? reportCode = null, string? sheetName = null)
        {
            if (path is null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            if (content is null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            // パスの正規化・出力先ディレクトリの作成で起きる例外も、書き込みの失敗として UtsushiException 階層へそろえる
            // (例: 「既存のファイル/out.pdf」への出力は Directory.CreateDirectory が IOException を投げる)。
            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(path);
                var directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory!);
                }
            }
            catch (Exception ex) when (IsPathOrIoFailure(ex))
            {
                throw new PdfRenderingException($"PDFの出力先を用意できません: {path}", reportCode, sheetName, ex);
            }

            // 同一ディレクトリ上の一時ファイルへ書いてから置き換える。
            // 書き込み途中で異常終了しても、出力先には旧ファイル(または何も)が残る。
            // ファイル名にGUIDを含めるのは、同じ出力パスへ同時に書き込む複数プロセスが
            // 同一の一時ファイルを取り合わないため(固定名だと、片方の書き込み完了直後に
            // もう片方が Exists→Delete→Move する間に正常な出力を消してしまうことがある)。
            var temporaryPath = fullPath + "." + Guid.NewGuid().ToString("N") + ".utsushi-tmp";
            try
            {
                using (var file = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    content.CopyTo(file);
                    file.Flush(flushToDisk: true);
                }

                // overwrite: true により、既存ファイルの有無を確認してから置き換えるまでの
                // TOCTOU(競合)の隙を無くす。
                File.Move(temporaryPath, fullPath, overwrite: true);
            }
            catch (Exception ex) when (IsPathOrIoFailure(ex))
            {
                TryDelete(temporaryPath);
                throw new PdfRenderingException($"PDFの書き出しに失敗しました: {fullPath}", reportCode, sheetName, ex);
            }
            catch
            {
                TryDelete(temporaryPath);
                throw;
            }
        }

        /// <summary>
        /// 出力先のパスやファイルシステムに起因する失敗かどうか。
        /// </summary>
        /// <remarks>
        /// <see cref="ArgumentException"/>(NUL文字などを含む不正なパス)・<see cref="NotSupportedException"/>
        /// (形式が不正なパス)も、呼び出し元が指定した出力先の問題として同じく扱う
        /// (null は呼び出し側の契約違反として事前に <see cref="ArgumentNullException"/> にしている)。
        /// </remarks>
        private static bool IsPathOrIoFailure(Exception ex) =>
            ex is IOException or UnauthorizedAccessException or NotSupportedException
            || (ex is ArgumentException && ex is not ArgumentNullException);

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // 後始末の失敗は元のエラーを覆い隠さないよう無視する。
            }
        }
    }
}
