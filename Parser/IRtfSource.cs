// -- FILE ------------------------------------------------------------------
// name       : IRtfSource.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.IO;

namespace Itenso.Rtf
{
    // ------------------------------------------------------------------------
    public interface IRtfSource
    {
        // ----------------------------------------------------------------------
        TextReader Reader { get; }
    } // interface IRtfSource
} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------