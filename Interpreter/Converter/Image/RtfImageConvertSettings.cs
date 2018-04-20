// -- FILE ------------------------------------------------------------------
// name       : RtfImageConvertSettings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.05
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using System.Drawing;
using System.IO;

namespace Itenso.Rtf.Converter.Image
{
    // ------------------------------------------------------------------------
    public class RtfImageConvertSettings
    {
        // ----------------------------------------------------------------------

        // ----------------------------------------------------------------------
        public IRtfVisualImageAdapter ImageAdapter { get; } // ImageAdapter

        // ----------------------------------------------------------------------
        public Color? BackgroundColor { get; set; }

        // ----------------------------------------------------------------------
        public string ImagesPath { get; set; } // ImagesPath

        // ----------------------------------------------------------------------
        public bool ScaleImage { get; set; } = true;

// ScaleImage

        // ----------------------------------------------------------------------
        public float ScaleOffset { get; set; }

        // ----------------------------------------------------------------------
        public float ScaleExtension { get; set; }

        // ----------------------------------------------------------------------
        public RtfImageConvertSettings() :
            this(new RtfVisualImageAdapter())
        {
        } // RtfImageConvertSettings

        // ----------------------------------------------------------------------
        public RtfImageConvertSettings(IRtfVisualImageAdapter imageAdapter)
        {
            if (imageAdapter == null)
                throw new ArgumentNullException("imageAdapter");

            ImageAdapter = imageAdapter;
        } // RtfImageConvertSettings

        // ----------------------------------------------------------------------
        public string GetImageFileName(int index, RtfVisualImageFormat rtfVisualImageFormat)
        {
            var imageFileName = ImageAdapter.ResolveFileName(index, rtfVisualImageFormat);
            if (!string.IsNullOrEmpty(ImagesPath))
                imageFileName = Path.Combine(ImagesPath, imageFileName);
            return imageFileName;
        } // GetImageFileName
    } // class RtfImageConvertSettings
} // namespace Itenso.Rtf.Converter.Image
// -- EOF -------------------------------------------------------------------