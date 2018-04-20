// name       : RtfHtmlConvertScope.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.09
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Rtf.Converter.Html
{
    [Flags]
    public enum RtfHtmlConvertScope
    {
        None = 0x00000000,

        Document = 0x00000001,
        Html = 0x00000010,
        Head = 0x00000100,
        Body = 0x00001000,
        Content = 0x00010000,

        All = Document | Html | Head | Body | Content
    }
}