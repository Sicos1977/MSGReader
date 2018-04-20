// name       : RtfDocument.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Rtf.Model
{
    public sealed class RtfDocument : IRtfDocument
    {
        // Members

        public RtfDocument(IRtfInterpreterContext context, IRtfVisualCollection visualContent) :
            this(context.RtfVersion,
                context.DefaultFont,
                context.FontTable,
                context.ColorTable,
                context.Generator,
                context.UniqueTextFormats,
                context.DocumentInfo,
                context.UserProperties,
                visualContent
            )
        {
        } // RtfDocument

        public RtfDocument(
            int rtfVersion,
            IRtfFont defaultFont,
            IRtfFontCollection fontTable,
            IRtfColorCollection colorTable,
            string generator,
            IRtfTextFormatCollection uniqueTextFormats,
            IRtfDocumentInfo documentInfo,
            IRtfDocumentPropertyCollection userProperties,
            IRtfVisualCollection visualContent
        )
        {
            if (rtfVersion != RtfSpec.RtfVersion1)
                throw new RtfUnsupportedStructureException(Strings.UnsupportedRtfVersion(rtfVersion));
            if (defaultFont == null)
                throw new ArgumentNullException(nameof(defaultFont));
            if (fontTable == null)
                throw new ArgumentNullException(nameof(fontTable));
            if (colorTable == null)
                throw new ArgumentNullException(nameof(colorTable));
            if (uniqueTextFormats == null)
                throw new ArgumentNullException(nameof(uniqueTextFormats));
            if (documentInfo == null)
                throw new ArgumentNullException(nameof(documentInfo));
            if (userProperties == null)
                throw new ArgumentNullException(nameof(userProperties));
            if (visualContent == null)
                throw new ArgumentNullException(nameof(visualContent));
            RtfVersion = rtfVersion;
            DefaultFont = defaultFont;
            DefaultTextFormat = new RtfTextFormat(defaultFont, RtfSpec.DefaultFontSize);
            FontTable = fontTable;
            ColorTable = colorTable;
            Generator = generator;
            UniqueTextFormats = uniqueTextFormats;
            DocumentInfo = documentInfo;
            UserProperties = userProperties;
            VisualContent = visualContent;
        } // RtfDocument

        public int RtfVersion { get; } // RtfVersion

        public IRtfFont DefaultFont { get; } // DefaultFont

        public IRtfTextFormat DefaultTextFormat { get; } // DefaultTextFormat

        public IRtfFontCollection FontTable { get; } // FontTable

        public IRtfColorCollection ColorTable { get; } // ColorTable

        public string Generator { get; } // Generator

        public IRtfTextFormatCollection UniqueTextFormats { get; } // UniqueTextFormats

        public IRtfDocumentInfo DocumentInfo { get; } // DocumentInfo

        public IRtfDocumentPropertyCollection UserProperties { get; } // UserProperties

        public IRtfVisualCollection VisualContent { get; } // VisualContent

        public override string ToString()
        {
            return "RTFv" + RtfVersion;
        } // ToString
    } // class RtfDocument
}