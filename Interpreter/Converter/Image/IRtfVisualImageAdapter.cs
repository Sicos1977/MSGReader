// -- FILE ------------------------------------------------------------------
// name       : IRtfVisualImageAdapter.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.05
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System.Drawing.Imaging;

namespace Itenso.Rtf.Converter.Image
{

	// ------------------------------------------------------------------------
	public interface IRtfVisualImageAdapter
	{

		// ----------------------------------------------------------------------
		string FileNamePattern { get; }

		// ----------------------------------------------------------------------
		ImageFormat TargetFormat { get; }

		// ----------------------------------------------------------------------
		double DpiX { get; }

		// ----------------------------------------------------------------------
		double DpiY { get; }

		// ----------------------------------------------------------------------
		ImageFormat GetImageFormat( RtfVisualImageFormat rtfVisualImageFormat );

		// ----------------------------------------------------------------------
		string ResolveFileName( int index, RtfVisualImageFormat rtfVisualImageFormat );

		// ----------------------------------------------------------------------
		int CalcImageWidth( RtfVisualImageFormat format, int width,
			int desiredWidth, int scaleWidthPercent );

		// ----------------------------------------------------------------------
		int CalcImageHeight( RtfVisualImageFormat format, int height,
			int desiredHeight, int scaleHeightPercent );

	} // interface IRtfVisualImageAdapter

} // namespace Itenso.Rtf.Converter.Image
// -- EOF -------------------------------------------------------------------
