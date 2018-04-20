// name       : RtfEmptyHtmlStyleConverter.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.07.13
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf.Converter.Html
{
    public class RtfEmptyHtmlStyleConverter : IRtfHtmlStyleConverter
    {
        public virtual IRtfHtmlStyle TextToHtml(IRtfVisualText visualText)
        {
            return RtfHtmlStyle.Empty;
        } // TextToHtml
    }
}