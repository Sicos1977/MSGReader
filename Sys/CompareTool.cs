// -- FILE ------------------------------------------------------------------
// name       : CompareTool.cs
// project    : System Framelet
// created    : Leon Poyyayil - 2005.05.12
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Sys
{
    // ------------------------------------------------------------------------
    public static class CompareTool
    {
        // ----------------------------------------------------------------------
        public static bool AreEqual(object left, object right)
        {
            return left == right || left != null && left.Equals(right);
        } // AreEqual
    } // class CompareTool
} // namespace Itenso.Sys
// -- EOF -------------------------------------------------------------------