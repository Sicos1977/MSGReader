// -- FILE ------------------------------------------------------------------
// name       : IRtfInterpreter.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfInterpreter
	{

		// ----------------------------------------------------------------------
		/// <summary>
		/// The settings used by the interpreter.
		/// </summary>
		IRtfInterpreterSettings Settings { get; }

		// ----------------------------------------------------------------------
		/// <summary>
		/// Adds a listener that will get notified along the interpretation process.
		/// </summary>
		/// <param name="listener">the listener to add</param>
		/// <exception cref="ArgumentNullException">in case of a null argument</exception>
		void AddInterpreterListener( IRtfInterpreterListener listener );

		// ----------------------------------------------------------------------
		/// <summary>
		/// Removes a listener from this instance.
		/// </summary>
		/// <param name="listener">the listener to remove</param>
		/// <exception cref="ArgumentNullException">in case of a null argument</exception>
		void RemoveInterpreterListener( IRtfInterpreterListener listener );

		// ----------------------------------------------------------------------
		/// <summary>
		/// Parses the given RTF document and informs the registered listeners about
		/// all occurring events.
		/// </summary>
		/// <param name="rtfDocument">the RTF documet to interpret</param>
		/// <exception cref="RtfException">in case of an unsupported RTF structure</exception>
		/// <exception cref="ArgumentNullException">in case of a null argument</exception>
		void Interpret( IRtfGroup rtfDocument );

	} // interface IRtfInterpreter

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
