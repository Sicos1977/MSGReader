// -- FILE ------------------------------------------------------------------
// name       : RtfParserTool.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System.IO;
using Itenso.Rtf.Parser;

namespace Itenso.Rtf.Support
{

	// ------------------------------------------------------------------------
	public static class RtfParserTool
	{

		// ----------------------------------------------------------------------
		public static IRtfGroup Parse( string rtfText, params IRtfParserListener[] listeners )
		{
			return Parse( new RtfSource( rtfText ), listeners );
		} // Parse

		// ----------------------------------------------------------------------
		public static IRtfGroup Parse( TextReader rtfTextSource, params IRtfParserListener[] listeners )
		{
			return Parse( new RtfSource( rtfTextSource ), listeners );
		} // Parse

		// ----------------------------------------------------------------------
		public static IRtfGroup Parse( Stream rtfTextSource, params IRtfParserListener[] listeners )
		{
			return Parse( new RtfSource( rtfTextSource ), listeners );
		} // Parse

		// ----------------------------------------------------------------------
		public static IRtfGroup Parse( IRtfSource rtfTextSource, params IRtfParserListener[] listeners )
		{
			RtfParserListenerStructureBuilder structureBuilder = new RtfParserListenerStructureBuilder();
			RtfParser parser = new RtfParser( structureBuilder );
			if ( listeners != null )
			{
				foreach ( IRtfParserListener listener in listeners )
				{
					if ( listener != null )
					{
						parser.AddParserListener( listener );
					}
				}
			}
			parser.Parse( rtfTextSource );
			return structureBuilder.StructureRoot;
		} // Parse

	} // class RtfParserTool

} // namespace Itenso.Rtf.Support
// -- EOF -------------------------------------------------------------------
