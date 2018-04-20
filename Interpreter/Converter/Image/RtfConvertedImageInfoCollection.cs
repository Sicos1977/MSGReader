// name       : RtfConvertedImageInfoCollection.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.10.13
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;
using Itenso.Sys.Collection;

namespace Itenso.Rtf.Converter.Image
{
    public sealed class RtfConvertedImageInfoCollection : ReadOnlyCollectionBase, IRtfConvertedImageInfoCollection
    {
        public IRtfConvertedImageInfo this[int index]
        {
            get { return InnerList[index] as RtfConvertedImageInfo; }
        } // this[ int ]

        public void CopyTo(IRtfConvertedImageInfo[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        public void Add(IRtfConvertedImageInfo item)
        {
            if (item == null)
                throw new ArgumentNullException("item");
            InnerList.Add(item);
        } // Add

        public void Clear()
        {
            InnerList.Clear();
        } // Clear

        public override string ToString()
        {
            return CollectionTool.ToString(this);
        } // ToString
    } // class RtfConvertedImageInfoCollection
}