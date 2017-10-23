// -- FILE ------------------------------------------------------------------
// name       : IRtfInterpreterListener.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfInterpreterListener
	{

		// ----------------------------------------------------------------------
		void BeginDocument( IRtfInterpreterContext context );

		// ----------------------------------------------------------------------
		void InsertText( IRtfInterpreterContext context, string text );

		// ----------------------------------------------------------------------
		void InsertSpecialChar( IRtfInterpreterContext context, RtfVisualSpecialCharKind kind );

		// ----------------------------------------------------------------------
		void InsertBreak( IRtfInterpreterContext context, RtfVisualBreakKind kind );

		// ----------------------------------------------------------------------
		void InsertImage( IRtfInterpreterContext context, RtfVisualImageFormat format,
			int width, int height, int desiredWidth, int desiredHeight,
			int scaleWidthPercent, int scaleHeightPercent, string imageDataHex
		);

		// ----------------------------------------------------------------------
		void EndDocument( IRtfInterpreterContext context );

	} // interface IRtfInterpreterListener

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
