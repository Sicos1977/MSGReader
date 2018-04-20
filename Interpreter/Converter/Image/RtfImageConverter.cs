// name       : RtfImageConverter.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.05.31
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using Itenso.Rtf.Interpreter;
using Itenso.Rtf.Model;

namespace Itenso.Rtf.Converter.Image
{
    public class RtfImageConverter : RtfInterpreterListenerBase
    {
        // members

        public RtfImageConvertSettings Settings { get; } // Settings

        public RtfConvertedImageInfoCollection ConvertedImages { get; } = new RtfConvertedImageInfoCollection();

// ConvertedImages

        public RtfImageConverter() :
            this(new RtfImageConvertSettings())
        {
        } // RtfImageConverter

        public RtfImageConverter(RtfImageConvertSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException("settings");

            Settings = settings;
        } // RtfImageConverter

        protected override void DoBeginDocument(IRtfInterpreterContext context)
        {
            base.DoBeginDocument(context);

            ConvertedImages.Clear();
        } // DoBeginDocument

        protected override void DoInsertImage(IRtfInterpreterContext context,
            RtfVisualImageFormat format,
            int width, int height,
            int desiredWidth, int desiredHeight,
            int scaleWidthPercent, int scaleHeightPercent,
            string imageDataHex
        )
        {
            var imageIndex = ConvertedImages.Count + 1;
            var fileName = Settings.GetImageFileName(imageIndex, format);
            EnsureImagesPath(fileName);

            var imageBuffer = RtfVisualImage.ToBinary(imageDataHex);
            Size imageSize;
            ImageFormat imageFormat;
            if (Settings.ImageAdapter.TargetFormat == null)
            {
                using (var image = System.Drawing.Image.FromStream(new MemoryStream(imageBuffer)))
                {
                    imageFormat = image.RawFormat;
                    imageSize = image.Size;
                }
                using (var binaryWriter = new BinaryWriter(File.Open(fileName, FileMode.Create)))
                {
                    binaryWriter.Write(imageBuffer);
                }
            }
            else
            {
                imageFormat = Settings.ImageAdapter.TargetFormat;
                if (Settings.ScaleImage)
                    imageSize = new Size(
                        Settings.ImageAdapter.CalcImageWidth(format, width, desiredWidth, scaleWidthPercent),
                        Settings.ImageAdapter.CalcImageHeight(format, height, desiredHeight, scaleHeightPercent));
                else
                    imageSize = new Size(width, height);

                SaveImage(imageBuffer, format, fileName, imageSize);
            }

            ConvertedImages.Add(new RtfConvertedImageInfo(fileName, imageFormat, imageSize));
        } // DoInsertImage

        protected virtual void SaveImage(byte[] imageBuffer, RtfVisualImageFormat format, string fileName, Size size)
        {
            var targetFormat = Settings.ImageAdapter.TargetFormat;

            var scaleOffset = Settings.ScaleOffset;
            var scaleExtension = Settings.ScaleExtension;
            using (var image = System.Drawing.Image.FromStream(
                new MemoryStream(imageBuffer, 0, imageBuffer.Length)))
            {
                var convertedImage = new Bitmap(new Bitmap(size.Width, size.Height, image.PixelFormat));
                var graphic = Graphics.FromImage(convertedImage);
                graphic.CompositingQuality = CompositingQuality.HighQuality;
                graphic.SmoothingMode = SmoothingMode.HighQuality;
                graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
                var rectangle = new RectangleF(
                    scaleOffset,
                    scaleOffset,
                    size.Width + scaleExtension,
                    size.Height + scaleExtension);

                if (Settings.BackgroundColor.HasValue)
                    graphic.Clear(Settings.BackgroundColor.Value);

                graphic.DrawImage(image, rectangle);
                convertedImage.Save(fileName, targetFormat);
            }
        } // SaveImage

        protected virtual void EnsureImagesPath(string imageFileName)
        {
            var fi = new FileInfo(imageFileName);
            if (!string.IsNullOrEmpty(fi.DirectoryName) && !Directory.Exists(fi.DirectoryName))
                Directory.CreateDirectory(fi.DirectoryName);
        } // EnsureImagesPath
    } // class RtfImageConverter
}