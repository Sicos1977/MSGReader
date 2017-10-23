// -- FILE ------------------------------------------------------------------
// name       : RtfParserListenerBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf.Parser
{

	// ------------------------------------------------------------------------
	public class RtfParserListenerBase : IRtfParserListener
	{

		// ----------------------------------------------------------------------
		public int Level
		{
			get { return level; }
		} // Level

		// ----------------------------------------------------------------------
		public void ParseBegin()
		{
			level = 0; // in case something interrupted the normal flow of things previously ...
			DoParseBegin();
		} // ParseBegin

		// ----------------------------------------------------------------------
		protected virtual void DoParseBegin()
		{
		} // DoParseBegin

		// ----------------------------------------------------------------------
		public void GroupBegin()
		{
			DoGroupBegin();
			level++;
		} // GroupBegin

		// ----------------------------------------------------------------------
		protected virtual void DoGroupBegin()
		{
		} // DoGroupBegin

		// ----------------------------------------------------------------------
		public void TagFound( IRtfTag tag )
		{
			if ( tag != null )
			{
				DoTagFound( tag );
			}
		} // TagFound

		// ----------------------------------------------------------------------
		protected virtual void DoTagFound( IRtfTag tag )
		{
		} // DoTagFound

		// ----------------------------------------------------------------------
		public void TextFound( IRtfText text )
		{
			if ( text != null )
			{
				DoTextFound( text );
			}
		} // TextFound

		// ----------------------------------------------------------------------
		protected virtual void DoTextFound( IRtfText text )
		{
		} // DoTextFound

		// ----------------------------------------------------------------------
		public void GroupEnd()
		{
			level--;
			DoGroupEnd();
		} // GroupEnd

		// ----------------------------------------------------------------------
		protected virtual void DoGroupEnd()
		{
		} // DoGroupEnd

		// ----------------------------------------------------------------------
		public void ParseSuccess()
		{
			DoParseSuccess();
		} // ParseSuccess

		// ----------------------------------------------------------------------
		protected virtual void DoParseSuccess()
		{
		} // DoParseSuccess

		// ----------------------------------------------------------------------
		public void ParseFail( RtfException reason )
		{
			DoParseFail( reason );
		} // ParseFail

		// ----------------------------------------------------------------------
		protected virtual void DoParseFail( RtfException reason )
		{
		} // DoParseFail

		// ----------------------------------------------------------------------
		public void ParseEnd()
		{
			DoParseEnd();
		} // ParseEnd

		// ----------------------------------------------------------------------
		protected virtual void DoParseEnd()
		{
		} // DoParseEnd

		// ----------------------------------------------------------------------
		// members
		private int level;

	} // class RtfParserListenerBase

} // namespace Itenso.Rtf.Parser
// -- EOF -------------------------------------------------------------------
