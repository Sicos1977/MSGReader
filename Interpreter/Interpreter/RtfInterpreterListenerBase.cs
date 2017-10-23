// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreterListenerBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf.Interpreter
{

	// ------------------------------------------------------------------------
	public class RtfInterpreterListenerBase : IRtfInterpreterListener
	{

		// ----------------------------------------------------------------------
		public void BeginDocument( IRtfInterpreterContext context )
		{
			if ( context != null )
			{
				DoBeginDocument( context );
			}
		} // BeginDocument

		// ----------------------------------------------------------------------
		public void InsertText( IRtfInterpreterContext context, string text )
		{
			if ( context != null )
			{
				DoInsertText( context, text );
			}
		} // InsertText

		// ----------------------------------------------------------------------
		public void InsertSpecialChar( IRtfInterpreterContext context, RtfVisualSpecialCharKind kind )
		{
			if ( context != null )
			{
				DoInsertSpecialChar( context, kind );
			}
		} // InsertSpecialChar

		// ----------------------------------------------------------------------
		public void InsertBreak( IRtfInterpreterContext context, RtfVisualBreakKind kind )
		{
			if ( context != null )
			{
				DoInsertBreak( context, kind );
			}
		} // InsertBreak

		// ----------------------------------------------------------------------
		public void InsertImage( IRtfInterpreterContext context, RtfVisualImageFormat format,
			int width, int height, int desiredWidth, int desiredHeight,
			int scaleWidthPercent, int scaleHeightPercent, string imageDataHex
		)
		{
			if ( context != null )
			{
				DoInsertImage( context, format,
					width, height, desiredWidth, desiredHeight,
					scaleWidthPercent, scaleHeightPercent, imageDataHex );
			}
		} // InsertImage

		// ----------------------------------------------------------------------
		public void EndDocument( IRtfInterpreterContext context )
		{
			if ( context != null )
			{
				DoEndDocument( context );
			}
		} // EndDocument

		// ----------------------------------------------------------------------
		protected virtual void DoBeginDocument( IRtfInterpreterContext context )
		{
		} // DoBeginDocument

		// ----------------------------------------------------------------------
		protected virtual void DoInsertText( IRtfInterpreterContext context, string text )
		{
		} // DoInsertText

		// ----------------------------------------------------------------------
		protected virtual void DoInsertSpecialChar( IRtfInterpreterContext context, RtfVisualSpecialCharKind kind )
		{
		} // DoInsertSpecialChar

		// ----------------------------------------------------------------------
		protected virtual void DoInsertBreak( IRtfInterpreterContext context, RtfVisualBreakKind kind )
		{
		} // DoInsertBreak

		// ----------------------------------------------------------------------
		protected virtual void DoInsertImage( IRtfInterpreterContext context,
			RtfVisualImageFormat format,
			int width, int height, int desiredWidth, int desiredHeight,
			int scaleWidthPercent, int scaleHeightPercent,
			string imageDataHex
		)
		{
		} // DoInsertImage

		// ----------------------------------------------------------------------
		protected virtual void DoEndDocument( IRtfInterpreterContext context )
		{
		} // DoEndDocument

	} // class RtfInterpreterListenerBase

} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------
