using System;
using System.IO;
using System.Text;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Xunit;

namespace Utsushi.Rendering.Tests
{
    /// <summary>
    /// <see cref="AtomicFileWriter"/> の検証(要件5.4)。
    /// </summary>
    /// <remarks>
    /// <see cref="SkiaPdfRenderer.RenderToFile"/> だけでなく、<c>Utsushi.ReportPdfConverter.ConvertToFile</c>
    /// が既定以外の <c>IPdfRenderer</c> を使う場合のフォールバック経路もこのヘルパーに委譲しているため、
    /// どちらの経路でも同じアトミック性が保証されることを確認する。
    /// </remarks>
    public sealed class AtomicFileWriterTests
    {
        [Fact]
        public void 内容を書き出し一時ファイルを残さない()
        {
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            var path = Path.Combine(directory, "out.txt");

            try
            {
                using (var content = new MemoryStream(Encoding.UTF8.GetBytes("hello")))
                {
                    AtomicFileWriter.Write(path, content);
                }

                Assert.True(File.Exists(path));
                Assert.Equal("hello", File.ReadAllText(path));
                Assert.False(File.Exists(path + ".utsushi-tmp"), "一時ファイルが残ってはいけない");
            }
            finally
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
        }

        [Fact]
        public void 存在しないディレクトリは作成される()
        {
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            var path = Path.Combine(directory, "nested", "out.txt");

            try
            {
                using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
                AtomicFileWriter.Write(path, content);

                Assert.True(File.Exists(path));
            }
            finally
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
        }

        [Fact]
        public void 既存のファイルを新しい内容で置き換える()
        {
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, "out.txt");
            File.WriteAllText(path, "旧い内容");

            try
            {
                using var content = new MemoryStream(Encoding.UTF8.GetBytes("新しい内容"));
                AtomicFileWriter.Write(path, content);

                Assert.Equal("新しい内容", File.ReadAllText(path));
            }
            finally
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
        }

        [Fact]
        public void 出力先の途中に既存のファイルがあればPdfRenderingExceptionになる()
        {
            // 「既存のファイル/out.pdf」への出力は、出力先ディレクトリの作成(Directory.CreateDirectory)が
            // IOException を投げる。以前は try の外だったため UtsushiException 階層の外へ漏れていた。
            var directory = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var existingFile = Path.Combine(directory, "existing.txt");
            File.WriteAllText(existingFile, "ファイル");

            try
            {
                using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));
                var ex = Assert.Throws<PdfRenderingException>(
                    () => AtomicFileWriter.Write(Path.Combine(existingFile, "out.pdf"), content, "invoice", "請求書"));

                Assert.IsAssignableFrom<IOException>(ex.InnerException);
                Assert.Equal(ProcessingStage.Rendering, ex.Stage);
                Assert.Equal("invoice", ex.ReportCode);
                Assert.Equal("請求書", ex.SheetName);
            }
            finally
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void NUL文字を含むパスはPdfRenderingExceptionになる()
        {
            using var content = new MemoryStream(Encoding.UTF8.GetBytes("hello"));

            var ex = Assert.Throws<PdfRenderingException>(
                () => AtomicFileWriter.Write(Path.Combine(Path.GetTempPath(), "a\0b.pdf"), content));

            Assert.IsAssignableFrom<ArgumentException>(ex.InnerException);
        }

        [Fact]
        public void パスがnullならArgumentNullExceptionになる()
        {
            using var content = new MemoryStream();

            Assert.Throws<ArgumentNullException>(() => AtomicFileWriter.Write(null!, content));
        }
    }
}
