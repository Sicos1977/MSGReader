// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreterException.cs
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
	public class RtfInterpreterException : RtfException
	{

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance.</summary>
		public RtfInterpreterException()
		{
		} // RtfInterpreterException

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance with the given message.</summary>
		/// <param name="message">the message to display</param>
		public RtfInterpreterException( string message ) :
			base( message )
		{
		} // RtfInterpreterException

		// ----------------------------------------------------------------------
		/// <summary>Creates a new instance with the given message, based on the given cause.</summary>
		/// <param name="message">the message to display</param>
		/// <param name="cause">the original cause for this exception</param>
		public RtfInterpreterException( string message, Exception cause ) :
			base( message, cause )
		{
		} // RtfInterpreterException

		// ----------------------------------------------------------------------
		/// <summary>Serialization support.</summary>
		/// <param name="info">the info to use for serialization</param>
		/// <param name="context">the context to use for serialization</param>
		protected RtfInterpreterException( SerializationInfo info, StreamingContext context ) :
			base( info, context )
		{
		} // RtfInterpreterException

	} // class RtfInterpreterException

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
