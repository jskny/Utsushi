using System;
using System.IO;
using Utsushi.Core.Exceptions;
using Utsushi.ReportDefinitions;
using Xunit;

namespace Utsushi.ReportDefinition.Tests
{
    /// <summary>
    /// <see cref="FileSystemReportDefinitionRepository"/> のパストラバーサル対策の検証。
    /// </summary>
    /// <remarks>
    /// 帳票コードは呼び出し元プロダクトから渡される外部入力(要件6.1)であるため、
    /// 不正な値でルートディレクトリの外を指せないことを確認する。
    /// </remarks>
    public sealed class FileSystemReportDefinitionRepositoryTests : IDisposable
    {
        private const string ValidJson = @"{
      ""schemaVersion"": 1,
      ""reportCode"": ""invoice"",
      ""sheetName"": ""請求書"",
      ""substitutionFields"": []
    }";

        private readonly string _root;

        public FileSystemReportDefinitionRepositoryTests()
        {
            _root = Path.Combine(Path.GetTempPath(), "utsushi-test-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_root);

            var reportDir = Path.Combine(_root, "invoice");
            Directory.CreateDirectory(reportDir);
            File.WriteAllText(
                Path.Combine(reportDir, FileSystemReportDefinitionRepository.DefinitionFileName), ValidJson);

            // ルートの外(親ディレクトリ)に、パストラバーサルが成功した場合にのみ読める定義を置く。
            File.WriteAllText(
                Path.Combine(_root, FileSystemReportDefinitionRepository.DefinitionFileName), ValidJson);
        }

        public void Dispose()
        {
            if (Directory.Exists(_root))
            {
                Directory.Delete(_root, recursive: true);
            }
        }

        [Theory]
        [InlineData("..")]
        [InlineData("../invoice")]
        [InlineData("../../etc/passwd")]
        [InlineData("..\\invoice")]
        [InlineData("foo/bar")]
        [InlineData("foo\\bar")]
        [InlineData(".")]
        public void パストラバーサルを試みる帳票コードは拒否される(string reportCode)
        {
            var repository = new FileSystemReportDefinitionRepository(_root);

            var ex = Assert.Throws<ReportDefinitionNotFoundException>(() => repository.Load(reportCode));
            Assert.Equal(reportCode, ex.ReportCode);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        public void 空の帳票コードは拒否される(string reportCode)
        {
            var repository = new FileSystemReportDefinitionRepository(_root);

            Assert.Throws<ReportDefinitionNotFoundException>(() => repository.Load(reportCode));
        }

        [Fact]
        public void Nullの帳票コードは拒否される()
        {
            var repository = new FileSystemReportDefinitionRepository(_root);

            Assert.Throws<ReportDefinitionNotFoundException>(() => repository.Load(null!));
        }

        [Fact]
        public void 正当な帳票コードは読み込める()
        {
            var repository = new FileSystemReportDefinitionRepository(_root);

            var definition = repository.Load("invoice");

            Assert.Equal("invoice", definition.ReportCode);
        }

        [Fact]
        public void 登録されていない帳票コードは見つからないエラーになる()
        {
            var repository = new FileSystemReportDefinitionRepository(_root);

            var ex = Assert.Throws<ReportDefinitionNotFoundException>(() => repository.Load("unknown"));
            Assert.Equal("unknown", ex.ReportCode);
        }
    }
}
