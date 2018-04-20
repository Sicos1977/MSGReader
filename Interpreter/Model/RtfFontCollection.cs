// -- FILE ------------------------------------------------------------------
// name       : RtfFontCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using System.Collections;

namespace Itenso.Rtf.Model
{
    // ------------------------------------------------------------------------
    public sealed class RtfFontCollection : ReadOnlyBaseCollection, IRtfFontCollection
    {
        // ----------------------------------------------------------------------
        // members
        private readonly Hashtable fontByIdMap = new Hashtable();

        // ----------------------------------------------------------------------
        public bool ContainsFontWithId(string fontId)
        {
            return fontByIdMap.ContainsKey(fontId);
        } // ContainsFontWithId

        // ----------------------------------------------------------------------
        public IRtfFont this[int index]
        {
            get { return InnerList[index] as IRtfFont; }
        } // this[ int ]

        // ----------------------------------------------------------------------
        public IRtfFont this[string id]
        {
            get { return fontByIdMap[id] as IRtfFont; }
        } // this[ string ]

        // ----------------------------------------------------------------------
        public void CopyTo(IRtfFont[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        // ----------------------------------------------------------------------
        public void Add(IRtfFont item)
        {
            if (item == null)
                throw new ArgumentNullException("item");
            InnerList.Add(item);
            fontByIdMap.Add(item.Id, item);
        } // Add

        // ----------------------------------------------------------------------
        public void Clear()
        {
            InnerList.Clear();
            fontByIdMap.Clear();
        } // Clear
    } // class RtfFontCollection
} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------