// name       : IRtfInterpreterSettings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2013.01.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf
{
    public interface IRtfInterpreterSettings
    {
        bool IgnoreDuplicatedFonts { get; set; }

        bool IgnoreUnknownFonts { get; set; }
    }
}