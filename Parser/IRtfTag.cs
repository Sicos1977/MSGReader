// -- FILE ------------------------------------------------------------------
// name       : IRtfTag.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfTag : IRtfElement
	{

		// ----------------------------------------------------------------------
		/// <summary>
		/// Returns the name together with the concatenated value as it stands in the rtf.
		/// </summary>
		string FullName { get; }

		// ----------------------------------------------------------------------
		string Name { get; }

		// ----------------------------------------------------------------------
		bool HasValue { get; }

		// ----------------------------------------------------------------------
		string ValueAsText { get; }

		// ----------------------------------------------------------------------
		int ValueAsNumber { get; }

	} // interface IRtfTag

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
