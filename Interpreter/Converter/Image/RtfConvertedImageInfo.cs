// name       : RtfConvertedImageInfo.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.10.13
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Drawing;
using System.Drawing.Imaging;

namespace Itenso.Rtf.Converter.Image
{
    public class RtfConvertedImageInfo : IRtfConvertedImageInfo
    {
        // members

        public RtfConvertedImageInfo(string fileName, ImageFormat format, Size size)
        {
            if (fileName == null)
                throw new ArgumentNullException("fileName");

            FileName = fileName;
            Format = format;
            Size = size;
        } // RtfConvertedImageInfo

        public string FileName { get; } // FileName

        public ImageFormat Format { get; } // Format

        public Size Size { get; } // Size

        public override string ToString()
        {
            return FileName + " " + Format + " " + Size.Width + "x" + Size.Height;
        } // ToString
    } // class RtfConvertedImageInfo
}