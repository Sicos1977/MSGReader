﻿// -- FILE ------------------------------------------------------------------
// name       : IRtfDocumentProperty.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfDocumentProperty
	{

		// ----------------------------------------------------------------------
		int PropertyKindCode { get; }

		// ----------------------------------------------------------------------
		RtfPropertyKind PropertyKind { get; }

		// ----------------------------------------------------------------------
		string Name { get; }

		// ----------------------------------------------------------------------
		string StaticValue { get; }

		// ----------------------------------------------------------------------
		string LinkValue { get; }

	} // interface IRtfDocumentProperty

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
