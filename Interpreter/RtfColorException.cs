// -- FILE ------------------------------------------------------------------
// name       : RtfColorException.cs
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
	public class RtfColorException : RtfInterpreterException
	{

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance.</summary>
		public RtfColorException()
		{
		} // RtfColorException

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance with the given message.</summary>
		/// <param name="message">the message to display</param>
		public RtfColorException( string message ) :
			base( message )
		{
		} // RtfColorException

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance with the given message, based on the given cause.</summary>
		/// <param name="message">the message to display</param>
		/// <param name="cause">the original cause for this exception</param>
		public RtfColorException( string message, Exception cause ) :
			base( message, cause )
		{
		} // RtfColorException

		// ----------------------------------------------------------------------
		/// <summary>Serialization support.</summary>
		/// <param name="info">the info to use for serialization</param>
		/// <param name="context">the context to use for serialization</param>
		protected RtfColorException( SerializationInfo info, StreamingContext context ) :
			base( info, context )
		{
		} // RtfColorException

	} // class RtfColorException

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
