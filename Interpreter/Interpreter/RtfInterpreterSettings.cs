// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreterSettings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2013.01.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf.Interpreter
{
    // ------------------------------------------------------------------------
    public sealed class RtfInterpreterSettings : IRtfInterpreterSettings
    {
        // ----------------------------------------------------------------------
        public bool IgnoreDuplicatedFonts { get; set; }

        // ----------------------------------------------------------------------
        public bool IgnoreUnknownFonts { get; set; }
    } // class RtfInterpreterSettings
} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------