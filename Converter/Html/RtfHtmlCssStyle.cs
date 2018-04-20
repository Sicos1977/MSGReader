// name       : RtfHtmlCssStyle.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.08
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections.Specialized;

namespace Itenso.Rtf.Converter.Html
{
    public class RtfHtmlCssStyle : IRtfHtmlCssStyle
    {
        // Members

        public RtfHtmlCssStyle(string selectorName)
        {
            if (selectorName == null)
                throw new ArgumentNullException(nameof(selectorName));
            SelectorName = selectorName;
        } // RtfHtmlCssStyle

        public NameValueCollection Properties { get; } = new NameValueCollection();

// Properties

        public string SelectorName { get; } // SelectorName
    }
}