// name       : RtfVisualCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Rtf.Model
{
    public sealed class RtfVisualCollection : ReadOnlyBaseCollection, IRtfVisualCollection
    {
        public IRtfVisual this[int index]
        {
            get { return InnerList[index] as IRtfVisual; }
        } // this[ int ]

        public void CopyTo(IRtfVisual[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        public void Add(IRtfVisual item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            InnerList.Add(item);
        } // Add

        public void Clear()
        {
            InnerList.Clear();
        } // Clear
    } // class RtfVisualCollection
}