using System.IO;
using System.Linq;
using Utsushi.Core;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// 出力しない画像を未対応要素として記録すること(excel2pdf-consolidation 要件2.5・2.7(d)(e))の検証。
    /// </summary>
    public sealed class UnsupportedImageRecordingTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        [Fact]
        public void 対応形式以外の画像を形式とアンカーとともに記録する()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/x-emf", ImageWorkbookFixtures.TinyPng());
            try
            {
                var workbook = _reader.ReadFile(path);

                var element = Assert.Single(workbook.UnsupportedElements);
                Assert.Equal(UnsupportedElementKind.UnsupportedImageFormat, element.Kind);
                Assert.Equal("テストシート", element.SheetName);
                Assert.Equal(CellAddress.Parse("B3"), element.Anchor);
                Assert.Equal("image/x-emf", element.Detail);
                Assert.False(workbook.UnsupportedElementsTruncated);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 内容が形式と一致しない画像を記録する()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", new byte[] { 0x00, 0x01, 0x02, 0x03 });
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.UnsupportedImageFormat, element.Kind);
                Assert.Equal("image/png", element.Detail);
                Assert.Contains("一致しない", element.Message);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 参照先の無い画像を記録する()
        {
            var path = ShapeWorkbookFixtures.CreateWorkbook(SafetyLimitWorkbookFixtures.PictureAnchorWithMissingEmbed());
            try
            {
                var element = Assert.Single(_reader.ReadFile(path).UnsupportedElements);

                Assert.Equal(UnsupportedElementKind.MissingImagePart, element.Kind);
                Assert.Equal("テストシート", element.SheetName);
                Assert.NotNull(element.Anchor);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像の枚数の上限の超過はシートごとに1件だけ記録する()
        {
            var path = ImageWorkbookFixtures.CreateWithManyPictures(OpenXmlWorkbookReader.MaxImagesPerSheet + 3);
            try
            {
                var workbook = _reader.ReadFile(path);

                Assert.Equal(OpenXmlWorkbookReader.MaxImagesPerSheet, workbook.Sheets[0].DrawingObjects.OfType<ImageModel>().Count());
                var element = Assert.Single(workbook.UnsupportedElements);
                Assert.Equal(UnsupportedElementKind.LimitExceeded, element.Kind);
                Assert.Equal("TooManyImages: 上限 " + OpenXmlWorkbookReader.MaxImagesPerSheet, element.Detail);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 正しい画像は記録しない()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng());
            try
            {
                Assert.Empty(_reader.ReadFile(path).UnsupportedElements);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
