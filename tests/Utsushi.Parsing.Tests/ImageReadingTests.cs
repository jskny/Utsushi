using System;
using System.IO;
using Utsushi.Core;
using Utsushi.Core.Exceptions;
using Utsushi.Parsing.Model;
using Utsushi.Parsing.OpenXml;
using Xunit;

namespace Utsushi.Parsing.Tests
{
    /// <summary>
    /// シート内画像(要件9)の読み取りの検証。
    /// </summary>
    public sealed class ImageReadingTests
    {
        private readonly OpenXmlWorkbookReader _reader = new();

        [Fact]
        public void oneCellAnchorの画像を固定サイズとして読み取る()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var image = Assert.Single(sheet.Images);

                Assert.Equal("image/png", image.ContentType);
                Assert.Equal(CellAddress.Parse("B3"), image.AnchorCell);
                Assert.Equal(0.0, image.AnchorOffset.X);
                Assert.Equal(0.0, image.AnchorOffset.Y);

                var extent = Assert.IsType<FixedImageExtent>(image.Extent);
                Assert.Equal(60.0, extent.WidthPt, 3);
                Assert.Equal(20.0, extent.HeightPt, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void twoCellAnchorの画像を対角セルとして読み取る()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng(), useTwoCellAnchor: true);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var image = Assert.Single(sheet.Images);

                Assert.Equal(CellAddress.Parse("B3"), image.AnchorCell);

                var extent = Assert.IsType<CellSpanImageExtent>(image.Extent);
                Assert.Equal(CellAddress.Parse("D5"), extent.ToCell);
                Assert.Equal(10.0, extent.ToOffset.X, 3);
                Assert.Equal(5.0, extent.ToOffset.Y, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 対応形式以外の画像はunsupportedElementsがignoreなら無視される()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/x-emf", ImageWorkbookFixtures.TinyPng());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.Images);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 対応形式以外の画像はunsupportedElementsがerrorなら例外になる()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/x-emf", ImageWorkbookFixtures.TinyPng());
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("UnsupportedImageFormat", ex.ElementKind);
                Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像のみのシートはunsupportedElementsがerrorでも例外にならない()
        {
            // DetectUnsupportedElementsが画像(xdr:pic)を「サポート外のDrawing」として
            // 誤検出しないことを確認する(design.md「Parsing レイヤー」の必須修正点)。
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng());
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var workbook = _reader.Read(stream, options);
                Assert.Single(workbook.Sheets[0].Images);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像以外の図形はunsupportedElementsがerrorなら引き続き例外になる()
        {
            var path = ImageWorkbookFixtures.CreateWithNonPictureShape();
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("Drawing", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ContentTypeを偽装した画像はunsupportedElementsがignoreなら無視される()
        {
            // ContentTypeはOPCパッケージ側の申告値に過ぎないため、実際のバイト列が
            // シグネチャと一致しない場合はunsupportedElementsポリシーに従う(security-reviewer指摘)。
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", new byte[] { 0x00, 0x01, 0x02, 0x03 });
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.Images);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ContentTypeを偽装した画像はunsupportedElementsがerrorなら例外になる()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", new byte[] { 0x00, 0x01, 0x02, 0x03 });
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("UnsupportedImageFormat", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像の数が上限を超える場合はunsupportedElementsがignoreなら上限までしか読み取らない()
        {
            var path = ImageWorkbookFixtures.CreateWithManyPictures(OpenXmlWorkbookReader.MaxImagesPerSheet + 5);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(OpenXmlWorkbookReader.MaxImagesPerSheet, sheet.Images.Count);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像の数が上限を超える場合はunsupportedElementsがerrorなら例外になる()
        {
            var path = ImageWorkbookFixtures.CreateWithManyPictures(OpenXmlWorkbookReader.MaxImagesPerSheet + 5);
            try
            {
                var options = new WorkbookReadOptions(UnsupportedElementBehavior.Error);
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.Read(stream, options));
                Assert.Equal("TooManyImages", ex.ElementKind);
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
