// name       : ArgumentCollection.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;

namespace Itenso.Sys.Application
{
    public sealed class ArgumentCollection : ReadOnlyCollectionBase
    {
        public IArgument this[int index]
        {
            get { return InnerList[index] as IArgument; }
        } // this[ int ]

        public bool IsValid
        {
            get
            {
                foreach (IArgument argument in InnerList)
                    if (!argument.IsValid)
                        return false;

                return true;
            }
        } // IsValid

        public void CopyTo(IArgument[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        public void Add(IArgument item)
        {
            if (item == null)
                throw new ArgumentNullException("item");
            InnerList.Add(item);
        } // Add

        public void Clear()
        {
            InnerList.Clear();
        } // Clear
    }
}