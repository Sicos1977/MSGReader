// -- FILE ------------------------------------------------------------------
// name       : RtfEncodingException.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2009.02.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Runtime.Serialization;

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	/// <summary>Thrown upon RTF specific error conditions.</summary>
	[Serializable]
	public class RtfEncodingException : RtfParserException
	{

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance.</summary>
		public RtfEncodingException()
		{
		} // RtfEncodingException

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance with the given message.</summary>
		/// <param name="message">the message to display</param>
		public RtfEncodingException( string message ) :
			base( message )
		{
		} // RtfEncodingException

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance with the given message, based on the given cause.</summary>
		/// <param name="message">the message to display</param>
		/// <param name="cause">the original cause for this exception</param>
		public RtfEncodingException( string message, Exception cause ) :
			base( message, cause )
		{
		} // RtfEncodingException

		// ----------------------------------------------------------------------
		/// <summary>Serialization support.</summary>
		/// <param name="info">the info to use for serialization</param>
		/// <param name="context">the context to use for serialization</param>
		protected RtfEncodingException( SerializationInfo info, StreamingContext context ) :
			base( info, context )
		{
		} // RtfEncodingException

	} // class RtfEncodingException

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
