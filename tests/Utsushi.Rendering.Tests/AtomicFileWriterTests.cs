using System;
using System.IO;
using System.Text;
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
    }
}
