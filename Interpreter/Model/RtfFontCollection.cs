// name       : RtfFontCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;

namespace Itenso.Rtf.Model
{
    public sealed class RtfFontCollection : ReadOnlyBaseCollection, IRtfFontCollection
    {
        // Members
        private readonly Hashtable _fontByIdMap = new Hashtable();

        public bool ContainsFontWithId(string fontId)
        {
            return _fontByIdMap.ContainsKey(fontId);
        } // ContainsFontWithId

        public IRtfFont this[int index]
        {
            get { return InnerList[index] as IRtfFont; }
        } // this[ int ]

        public IRtfFont this[string id]
        {
            get { return _fontByIdMap[id] as IRtfFont; }
        } // this[ string ]

        public void CopyTo(IRtfFont[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        public void Add(IRtfFont item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            InnerList.Add(item);
            _fontByIdMap.Add(item.Id, item);
        } // Add

        public void Clear()
        {
            InnerList.Clear();
            _fontByIdMap.Clear();
        } // Clear
    } // class RtfFontCollection
}