using System;
using System.IO;
using Utsushi.Core.Exceptions;

namespace Utsushi.Rendering;

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
    /// <exception cref="PdfRenderingException">書き込みに失敗した場合。</exception>
    public static void Write(string path, Stream content, string? reportCode = null, string? sheetName = null)
    {
        var fullPath = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory!);
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
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
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
