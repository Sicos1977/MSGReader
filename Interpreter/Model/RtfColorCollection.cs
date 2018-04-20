// -- FILE ------------------------------------------------------------------
// name       : RtfColorCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;

namespace Itenso.Rtf.Model
{
    // ------------------------------------------------------------------------
    public sealed class RtfColorCollection : ReadOnlyBaseCollection, IRtfColorCollection
    {
        // ----------------------------------------------------------------------
        public IRtfColor this[int index]
        {
            get { return InnerList[index] as IRtfColor; }
        } // this[ int ]

        // ----------------------------------------------------------------------
        public void CopyTo(IRtfColor[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        // ----------------------------------------------------------------------
        public void Add(IRtfColor item)
        {
            if (item == null)
                throw new ArgumentNullException("item");
            InnerList.Add(item);
        } // Add

        // ----------------------------------------------------------------------
        public void Clear()
        {
            InnerList.Clear();
        } // Clear
    } // class RtfColorCollection
} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------