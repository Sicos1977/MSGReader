// name       : RtfHtmlCssStyleCollection.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.08
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;

namespace Itenso.Rtf.Converter.Html
{
    public sealed class RtfHtmlCssStyleCollection : ReadOnlyCollectionBase, IRtfHtmlCssStyleCollection
    {
        public IRtfHtmlCssStyle this[int index] => InnerList[index] as RtfHtmlCssStyle;
        // this[ int ]

        public bool Contains(string selectorName)
        {
            foreach (IRtfHtmlCssStyle cssStyle in InnerList)
                if (cssStyle.SelectorName.Equals(selectorName))
                    return true;
            return false;
        } // Contains

        public void CopyTo(IRtfHtmlCssStyle[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        public void Add(IRtfHtmlCssStyle item)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(item));
            InnerList.Add(item);
        } // Add

        public void Clear()
        {
            InnerList.Clear();
        } // Clear
    }
}