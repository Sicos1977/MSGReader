﻿// -- FILE ------------------------------------------------------------------
// name       : IRtfVisualText.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.22
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{

	// ------------------------------------------------------------------------
	public interface IRtfVisualText : IRtfVisual
	{

		// ----------------------------------------------------------------------
		string Text { get; }

		// ----------------------------------------------------------------------
		IRtfTextFormat Format { get; }

	} // interface IRtfVisualText

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
