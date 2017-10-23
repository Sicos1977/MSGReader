// -- FILE ------------------------------------------------------------------
// name       : RtfVisualVisitorBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.26
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf.Support
{

	// ------------------------------------------------------------------------
	public class RtfVisualVisitorBase : IRtfVisualVisitor
	{

		// ----------------------------------------------------------------------
		public void VisitText( IRtfVisualText visualText )
		{
			if ( visualText != null )
			{
				DoVisitText( visualText );
			}
		} // VisitText

		// ----------------------------------------------------------------------
		protected virtual void DoVisitText( IRtfVisualText visualText )
		{
		} // DoVisitText

		// ----------------------------------------------------------------------
		public void VisitBreak( IRtfVisualBreak visualBreak )
		{
			if ( visualBreak != null )
			{
				DoVisitBreak( visualBreak );
			}
		} // VisitBreak

		// ----------------------------------------------------------------------
		protected virtual void DoVisitBreak( IRtfVisualBreak visualBreak )
		{
		} // DoVisitBreak

		// ----------------------------------------------------------------------
		public void VisitSpecial( IRtfVisualSpecialChar visualSpecialChar )
		{
			if ( visualSpecialChar != null )
			{
				DoVisitSpecial( visualSpecialChar );
			}
		} // VisitSpecial

		// ----------------------------------------------------------------------
		protected virtual void DoVisitSpecial( IRtfVisualSpecialChar visualSpecialChar )
		{
		} // DoVisitSpecial

		// ----------------------------------------------------------------------
		public void VisitImage( IRtfVisualImage visualImage )
		{
			if ( visualImage != null )
			{
				DoVisitImage( visualImage );
			}
		} // VisitImage

		// ----------------------------------------------------------------------
		protected virtual void DoVisitImage( IRtfVisualImage visualImage )
		{
		} // DoVisitImage

	} // class RtfVisualVisitorBase

} // namespace Itenso.Rtf.Support
// -- EOF -------------------------------------------------------------------
