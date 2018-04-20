// -- FILE ------------------------------------------------------------------
// name       : IRtfConvertedImageInfo.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.10.13
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Drawing;
using System.Drawing.Imaging;

namespace Itenso.Rtf.Converter.Image
{
    // ------------------------------------------------------------------------
    public interface IRtfConvertedImageInfo
    {
        // ----------------------------------------------------------------------
        string FileName { get; }

        // ----------------------------------------------------------------------
        ImageFormat Format { get; }

        // ----------------------------------------------------------------------
        Size Size { get; }
    } // interface IRtfConvertedImageInfo
} // namespace Itenso.Rtf.Converter.Image
// -- EOF -------------------------------------------------------------------