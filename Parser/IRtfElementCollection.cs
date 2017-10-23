// -- FILE ------------------------------------------------------------------
// name       : IRtfElementCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System.Collections;

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfElementCollection : IEnumerable
	{

		// ----------------------------------------------------------------------
		int Count { get; }

		// ----------------------------------------------------------------------
		IRtfElement this[ int index ] { get; }

		// ----------------------------------------------------------------------
		void CopyTo( IRtfElement[] array, int index );

	} // interface IRtfElementCollection

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
