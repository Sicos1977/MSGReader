// name       : IRtfDocument.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf
{
    public interface IRtfDocument
    {
        int RtfVersion { get; }

        IRtfFont DefaultFont { get; }

        IRtfTextFormat DefaultTextFormat { get; }

        IRtfFontCollection FontTable { get; }

        IRtfColorCollection ColorTable { get; }

        string Generator { get; }

        IRtfTextFormatCollection UniqueTextFormats { get; }

        IRtfDocumentInfo DocumentInfo { get; }

        IRtfDocumentPropertyCollection UserProperties { get; }

        IRtfVisualCollection VisualContent { get; }
    }
}