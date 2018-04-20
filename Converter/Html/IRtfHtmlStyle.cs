// name       : IRtfHtmlStyle.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.08
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf.Converter.Html
{
    public interface IRtfHtmlStyle
    {
        string ForegroundColor { get; set; }

        string BackgroundColor { get; set; }

        string FontFamily { get; set; }

        string FontSize { get; set; }

        bool IsEmpty { get; }
    }
}