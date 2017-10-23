// -- FILE ------------------------------------------------------------------
// name       : IRtfColorCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System.Collections;

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfColorCollection : IEnumerable
	{

		// ----------------------------------------------------------------------
		int Count { get; }

		// ----------------------------------------------------------------------
		IRtfColor this[ int index ] { get; }

		// ----------------------------------------------------------------------
		void CopyTo( IRtfColor[] array, int index );

	} // interface IRtfColorCollection

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
