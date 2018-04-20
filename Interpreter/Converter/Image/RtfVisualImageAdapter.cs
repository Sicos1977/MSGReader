// -- FILE ------------------------------------------------------------------
// name       : RtfVisualImageAdapter.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.05
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using System.Drawing.Imaging;
using System.Globalization;

namespace Itenso.Rtf.Converter.Image
{
    // ------------------------------------------------------------------------
    public class RtfVisualImageAdapter : IRtfVisualImageAdapter
    {
        // ----------------------------------------------------------------------
        public const double DefaultDpi = 96.0;

        private const string defaultFileNamePattern = "{0}{1}";
        private const int twipsPerInch = 1440;

        // ----------------------------------------------------------------------
        // members

        // ----------------------------------------------------------------------
        public RtfVisualImageAdapter() :
            this(defaultFileNamePattern, null)
        {
        } // RtfVisualImageAdapter

        // ----------------------------------------------------------------------
        public RtfVisualImageAdapter(string fileNamePattern) :
            this(fileNamePattern, null)
        {
        } // RtfVisualImageAdapter

        // ----------------------------------------------------------------------
        public RtfVisualImageAdapter(ImageFormat targetFormat) :
            this(defaultFileNamePattern, targetFormat)
        {
        } // RtfVisualImageAdapter

        // ----------------------------------------------------------------------
        public RtfVisualImageAdapter(string fileNamePattern, ImageFormat targetFormat) :
            this(fileNamePattern, targetFormat, DefaultDpi, DefaultDpi)
        {
            if (fileNamePattern == null)
                throw new ArgumentNullException("fileNamePattern");

            FileNamePattern = fileNamePattern;
            TargetFormat = targetFormat;
        } // RtfVisualImageAdapter

        // ----------------------------------------------------------------------
        public RtfVisualImageAdapter(string fileNamePattern, ImageFormat targetFormat, double dpiX, double dpiY)
        {
            if (fileNamePattern == null)
                throw new ArgumentNullException("fileNamePattern");

            FileNamePattern = fileNamePattern;
            TargetFormat = targetFormat;
            DpiX = dpiX;
            DpiY = dpiY;
        } // RtfVisualImageAdapter

        // ----------------------------------------------------------------------
        public string FileNamePattern { get; } // FileNamePattern

        // ----------------------------------------------------------------------
        public ImageFormat TargetFormat { get; } // TargetFormat

        // ----------------------------------------------------------------------
        public double DpiX { get; } // DpiX

        // ----------------------------------------------------------------------
        public double DpiY { get; } // DpiY

        // ----------------------------------------------------------------------
        public ImageFormat GetImageFormat(RtfVisualImageFormat rtfVisualImageFormat)
        {
            ImageFormat imageFormat = null;

            switch (rtfVisualImageFormat)
            {
                case RtfVisualImageFormat.Emf:
                    imageFormat = ImageFormat.Emf;
                    break;
                case RtfVisualImageFormat.Png:
                    imageFormat = ImageFormat.Png;
                    break;
                case RtfVisualImageFormat.Jpg:
                    imageFormat = ImageFormat.Jpeg;
                    break;
                case RtfVisualImageFormat.Wmf:
                    imageFormat = ImageFormat.Wmf;
                    break;
                case RtfVisualImageFormat.Bmp:
                    imageFormat = ImageFormat.Bmp;
                    break;
            }

            return imageFormat;
        } // GetImageFormat

        // ----------------------------------------------------------------------
        public string ResolveFileName(int index, RtfVisualImageFormat rtfVisualImageFormat)
        {
            var imageFormat = TargetFormat ?? GetImageFormat(rtfVisualImageFormat);

            return string.Format(
                CultureInfo.InvariantCulture,
                FileNamePattern,
                index,
                GetFileImageExtension(imageFormat));
        } // ResolveFileName

        // ----------------------------------------------------------------------
        public int CalcImageWidth(RtfVisualImageFormat format, int width,
            int desiredWidth, int scaleWidthPercent)
        {
            var imgScaleX = scaleWidthPercent / 100.0f;
            return (int) Math.Round((double) desiredWidth * imgScaleX / twipsPerInch * DpiX);
        } // CalcImageWidth

        // ----------------------------------------------------------------------
        public int CalcImageHeight(RtfVisualImageFormat format, int height,
            int desiredHeight, int scaleHeightPercent)
        {
            var imgScaleY = scaleHeightPercent / 100.0f;
            return (int) Math.Round((double) desiredHeight * imgScaleY / twipsPerInch * DpiY);
        } // CalcImageHeight

        // ----------------------------------------------------------------------
        private static string GetFileImageExtension(ImageFormat imageFormat)
        {
            string imageExtension = null;

            if (imageFormat == ImageFormat.Bmp)
                imageExtension = ".bmp";
            else if (imageFormat == ImageFormat.Emf)
                imageExtension = ".emf";
            else if (imageFormat == ImageFormat.Exif)
                imageExtension = ".exif";
            else if (imageFormat == ImageFormat.Gif)
                imageExtension = ".gif";
            else if (imageFormat == ImageFormat.Icon)
                imageExtension = ".ico";
            else if (imageFormat == ImageFormat.Jpeg)
                imageExtension = ".jpg";
            else if (imageFormat == ImageFormat.Png)
                imageExtension = ".png";
            else if (imageFormat == ImageFormat.Tiff)
                imageExtension = ".tiff";
            else if (imageFormat == ImageFormat.Wmf)
                imageExtension = ".wmf";

            return imageExtension;
        } // GetFileImageExtension
    } // class RtfVisualImageAdapter
} // namespace Itenso.Rtf.Converter.Image
// -- EOF -------------------------------------------------------------------