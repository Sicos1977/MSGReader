// -- FILE ------------------------------------------------------------------
// name       : IRtfElement.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfElement
	{

		// ----------------------------------------------------------------------
		RtfElementKind Kind { get; }

		// ----------------------------------------------------------------------
		void Visit( IRtfElementVisitor visitor );

	} // interface IRtfElement

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
