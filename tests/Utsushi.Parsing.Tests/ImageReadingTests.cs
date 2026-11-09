using System;
using System.IO;
using System.Linq;
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
                var image = Assert.Single(sheet.DrawingObjects.OfType<ImageModel>().ToList());

                Assert.Equal("image/png", image.ContentType);
                Assert.Equal(CellAddress.Parse("B3"), image.AnchorCell);
                Assert.Equal(0.0, image.AnchorOffset.X);
                Assert.Equal(0.0, image.AnchorOffset.Y);

                var extent = Assert.IsType<FixedAnchorExtent>(image.Extent);
                Assert.Equal(60.0, extent.WidthPt, 3);
                Assert.Equal(20.0, extent.HeightPt, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像のIdはNonVisualDrawingProperties_idを反映する()
        {
            // 要件10.11: 接続線の接続先解決のキーとなるIdを画像側でも保持する。
            // ImageWorkbookFixtures.CreateWithPicture(oneCellAnchor)はNonVisualDrawingProperties.Idを
            // 固定値2で組み立てる(BuildPictureの既定引数)。
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var image = Assert.Single(sheet.DrawingObjects.OfType<ImageModel>().ToList());

                Assert.Equal(2u, image.Id);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 回転角を度に変換して読み取る()
        {
            // 60,000分の1度単位。45度 = 2,700,000(要件9.7。捺印画像等の回転配置)。
            var path = ImageWorkbookFixtures.CreateWithPicture(
                "image/png", ImageWorkbookFixtures.TinyPng(), rotationEmu: 2_700_000);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var image = Assert.Single(sheet.DrawingObjects.OfType<ImageModel>().ToList());

                Assert.Equal(45.0, image.RotationDegrees, 3);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 回転の指定が無い画像は回転角0として読み取る()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var image = Assert.Single(sheet.DrawingObjects.OfType<ImageModel>().ToList());

                Assert.Equal(0.0, image.RotationDegrees, 3);
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
                var image = Assert.Single(sheet.DrawingObjects.OfType<ImageModel>().ToList());

                Assert.Equal(CellAddress.Parse("B3"), image.AnchorCell);

                var extent = Assert.IsType<CellSpanAnchorExtent>(image.Extent);
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
        public void 対応形式以外の画像は読み込みでは無視される()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/x-emf", ImageWorkbookFixtures.TinyPng());
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects.OfType<ImageModel>().ToList());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 対応形式以外の画像は未対応要素として記録する()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/x-emf", ImageWorkbookFixtures.TinyPng());
            try
            {
                var options = WorkbookReadOptions.Default;
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.ReadStrict(stream, options));
                StrictReading.AssertContains(ex, UnsupportedElementKind.UnsupportedImageFormat);
                Assert.Equal(ProcessingStage.Parsing, ex.Stage);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像のみのシートは未対応要素として記録しない()
        {
            // 以前の DetectUnsupportedElements が画像(xdr:pic)を「対応していない Drawing」として
            // 誤検出しないことを確認する(design.md「Parsing レイヤー」の必須修正点)。
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", ImageWorkbookFixtures.TinyPng());
            try
            {
                var options = WorkbookReadOptions.Default;
                using var stream = File.OpenRead(path);

                var workbook = _reader.ReadStrict(stream, options);
                Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<ImageModel>().ToList());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 対応済みプリセットの接続線は未対応要素として記録しない()
        {
            // 要件10.9(接続線対応)により、straightConnector1等の対応済みプリセットは
            // 構造的な未対応要素(旧: ElementKind="Drawing")の対象から外れ、ConnectorModelとして読み取られる。
            var path = ImageWorkbookFixtures.CreateWithConnectionShape();
            try
            {
                var options = WorkbookReadOptions.Default;
                using var stream = File.OpenRead(path);

                var workbook = _reader.ReadStrict(stream, options);
                var connector = Assert.Single(workbook.Sheets[0].DrawingObjects.OfType<ConnectorModel>().ToList());
                Assert.Equal(ConnectorPresetType.Straight, connector.Preset);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ContentTypeを偽装した画像は読み込みでは無視される()
        {
            // ContentTypeはOPCパッケージ側の申告値に過ぎないため、実際のバイト列が
            // シグネチャと一致しない場合はunsupportedElementsポリシーに従う(security-reviewer指摘)。
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", new byte[] { 0x00, 0x01, 0x02, 0x03 });
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects.OfType<ImageModel>().ToList());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void ContentTypeを偽装した画像は未対応要素として記録する()
        {
            var path = ImageWorkbookFixtures.CreateWithPicture("image/png", new byte[] { 0x00, 0x01, 0x02, 0x03 });
            try
            {
                var options = WorkbookReadOptions.Default;
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.ReadStrict(stream, options));
                StrictReading.AssertContains(ex, UnsupportedElementKind.UnsupportedImageFormat);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像の数が上限を超える場合は読み込みでは上限までしか読み取らない()
        {
            var path = ImageWorkbookFixtures.CreateWithManyPictures(OpenXmlWorkbookReader.MaxImagesPerSheet + 5);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Equal(OpenXmlWorkbookReader.MaxImagesPerSheet, sheet.DrawingObjects.OfType<ImageModel>().ToList().Count);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 画像の数が上限を超える場合は未対応要素として記録する()
        {
            var path = ImageWorkbookFixtures.CreateWithManyPictures(OpenXmlWorkbookReader.MaxImagesPerSheet + 5);
            try
            {
                var options = WorkbookReadOptions.Default;
                using var stream = File.OpenRead(path);

                var ex = Assert.Throws<UnsupportedWorkbookElementException>(() => _reader.ReadStrict(stream, options));
                StrictReading.AssertContains(ex, UnsupportedElementKind.LimitExceeded, "TooManyImages");
            }
            finally
            {
                File.Delete(path);
            }
        }

        // --- 画像の上限は、検証の成否によらず読み取りを試みた数で数える(security-reviewer指摘) ----
        // 検証に失敗した画像を数えないと、同じ画像パートを参照する壊れたアンカーを大量に並べて、
        // 上限に達しないまま画像を何度も展開させられる。

        [Fact]
        public void 検証に失敗する画像が上限より1少ない後ろの正しい画像は読み取られる()
        {
            // 上限-1 個の不正な画像 + 正しい画像 = ちょうど上限枚の読み取りを試みる。
            var path = ImageWorkbookFixtures.CreateWithInvalidPicturesThenValid(OpenXmlWorkbookReader.MaxImagesPerSheet - 1);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                var image = Assert.Single(sheet.DrawingObjects.OfType<ImageModel>().ToList());
                Assert.Equal((uint)(OpenXmlWorkbookReader.MaxImagesPerSheet - 1 + 100), image.Id);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 検証に失敗する画像が上限個並ぶと後ろの正しい画像は読み取られない()
        {
            // 上限+1 枚目の読み取りは試みない。検証に失敗した画像も数えていることの確認。
            var path = ImageWorkbookFixtures.CreateWithInvalidPicturesThenValid(OpenXmlWorkbookReader.MaxImagesPerSheet);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects.OfType<ImageModel>().ToList());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void 検証に失敗する画像が上限個並ぶと要素ごとに記録し後ろの正しい画像は上限の超過として1件記録する()
        {
            // 検証に失敗した画像も枚数に数える(上限の数え方)。不正な画像は1枚ずつ記録し、上限を超えた後ろの正しい画像は
            // 枚数の上限の超過として1件だけ記録する(excel2pdf-consolidation 要件2.5)。
            var path = ImageWorkbookFixtures.CreateWithInvalidPicturesThenValid(OpenXmlWorkbookReader.MaxImagesPerSheet);
            try
            {
                var elements = _reader.ReadFile(path).UnsupportedElements;

                Assert.Equal(OpenXmlWorkbookReader.MaxImagesPerSheet, elements.Count(e => e.Kind == UnsupportedElementKind.UnsupportedImageFormat));
                var limit = Assert.Single(elements, e => e.Kind == UnsupportedElementKind.LimitExceeded);
                Assert.StartsWith("TooManyImages", limit.Detail);
                Assert.Equal(OpenXmlWorkbookReader.MaxImagesPerSheet + 1, elements.Count);
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内で検証に失敗する画像が上限より1少ない後ろの正しい画像は読み取られる()
        {
            var path = ImageWorkbookFixtures.CreateWithInvalidPicturesThenValid(
                OpenXmlWorkbookReader.MaxImagesPerSheet - 1, invalidInGroups: true);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);

                // 不正な画像を含むグループは 読み込みではグループごと破棄される。
                Assert.Empty(sheet.DrawingObjects.OfType<GroupShapeModel>().ToList());
                Assert.Single(sheet.DrawingObjects.OfType<ImageModel>().ToList());
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void グループ内で検証に失敗する画像も数えるため上限個並ぶと後ろの正しい画像は読み取られない()
        {
            var path = ImageWorkbookFixtures.CreateWithInvalidPicturesThenValid(
                OpenXmlWorkbookReader.MaxImagesPerSheet, invalidInGroups: true);
            try
            {
                var sheet = Assert.Single(_reader.ReadFile(path).Sheets);
                Assert.Empty(sheet.DrawingObjects.OfType<ImageModel>().ToList());
            }
            finally
            {
                File.Delete(path);
            }
        }
    }
}
