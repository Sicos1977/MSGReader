// name       : RtfDocumentPropertyCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Rtf.Model
{
    public sealed class RtfDocumentPropertyCollection : ReadOnlyBaseCollection, IRtfDocumentPropertyCollection
    {
        public IRtfDocumentProperty this[int index]
        {
            get { return InnerList[index] as IRtfDocumentProperty; }
        } // this[ int ]

        public IRtfDocumentProperty this[string name]
        {
            get
            {
                if (name != null)
                    foreach (IRtfDocumentProperty property in InnerList)
                        if (property.Name.Equals(name))
                            return property;
                return null;
            }
        } // this[ string ]

        public void CopyTo(IRtfDocumentProperty[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        public void Add(IRtfDocumentProperty item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            InnerList.Add(item);
        } // Add

        public void Clear()
        {
            InnerList.Clear();
        } // Clear
    } // class RtfDocumentPropertyCollection
}