// -- FILE ------------------------------------------------------------------
// name       : RtfSource.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.IO;

namespace Itenso.Rtf.Support
{

	// ------------------------------------------------------------------------
	public sealed class RtfSource : IRtfSource
	{

		// ----------------------------------------------------------------------
		public RtfSource( string rtf )
		{
			if ( rtf == null )
			{
				throw new ArgumentNullException( "rtf" );
			}
			reader = new StringReader( rtf );
		} // RtfSource

		// ----------------------------------------------------------------------
		public RtfSource( TextReader rtf )
		{
			if ( rtf == null )
			{
				throw new ArgumentNullException( "rtf" );
			}
			reader = rtf;
		} // RtfSource

		// ----------------------------------------------------------------------
		public RtfSource( Stream rtf )
		{
			if ( rtf == null )
			{
				throw new ArgumentNullException( "rtf" );
			}
			reader = new StreamReader( rtf, RtfSpec.AnsiEncoding );
		} // RtfSource

		// ----------------------------------------------------------------------
		public TextReader Reader
		{
			get { return reader; }
		} // Reader

		// ----------------------------------------------------------------------
		// members
		private readonly TextReader reader;

	} // class RtfSource

} // namespace Itenso.Rtf.Support
// -- EOF -------------------------------------------------------------------
