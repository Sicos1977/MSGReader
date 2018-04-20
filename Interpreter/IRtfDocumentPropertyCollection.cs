// -- FILE ------------------------------------------------------------------
// name       : IRtfDocumentPropertyCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Collections;

namespace Itenso.Rtf
{
    // ------------------------------------------------------------------------
    public interface IRtfDocumentPropertyCollection : IEnumerable
    {
        // ----------------------------------------------------------------------
        int Count { get; }

        // ----------------------------------------------------------------------
        IRtfDocumentProperty this[int index] { get; }

        // ----------------------------------------------------------------------
        IRtfDocumentProperty this[string name] { get; }

        // ----------------------------------------------------------------------
        void CopyTo(IRtfDocumentProperty[] array, int index);
    } // interface IRtfDocumentPropertyCollection
} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------