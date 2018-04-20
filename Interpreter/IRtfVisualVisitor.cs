// -- FILE ------------------------------------------------------------------
// name       : IRtfVisualVisitor.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.26
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{
    // ------------------------------------------------------------------------
    public interface IRtfVisualVisitor
    {
        // ----------------------------------------------------------------------
        void VisitText(IRtfVisualText visualText);

        // ----------------------------------------------------------------------
        void VisitBreak(IRtfVisualBreak visualBreak);

        // ----------------------------------------------------------------------
        void VisitSpecial(IRtfVisualSpecialChar visualSpecialChar);

        // ----------------------------------------------------------------------
        void VisitImage(IRtfVisualImage visualImage);
    } // interface IRtfVisualVisitor
} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------