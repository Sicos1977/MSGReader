// name       : IRtfHtmlCssStyle.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.08
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Collections.Specialized;

namespace Itenso.Rtf.Converter.Html
{
    public interface IRtfHtmlCssStyle
    {
        NameValueCollection Properties { get; }

        string SelectorName { get; }
    }
}