// name       : RtfFontBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Text;
using Itenso.Rtf.Model;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    public sealed class RtfFontBuilder : RtfElementVisitorBase
    {
        private readonly StringBuilder _fontNameBuffer = new StringBuilder();

        // Members

        public string FontId { get; private set; } // FontId

        public int FontIndex { get; private set; } // FontIndex

        public int FontCharset { get; private set; } // FontCharset

        public int FontCodePage { get; private set; } // FontCodePage

        public RtfFontKind FontKind { get; private set; } // FontKind

        public RtfFontPitch FontPitch { get; private set; } // FontPitch

        public string FontName
        {
            get
            {
                string fontName = null;
                var len = _fontNameBuffer.Length;
                if (len > 0 && _fontNameBuffer[len - 1] == ';')
                {
                    fontName = _fontNameBuffer.ToString().Substring(0, len - 1).Trim();
                    if (fontName.Length == 0)
                        fontName = null;
                }
                return fontName;
            }
        } // FontName

        public RtfFontBuilder() :
            base(RtfElementVisitorOrder.NonRecursive)
        {
            // we iterate over our children ourselves -> hence non-recursive
            Reset();
        } // RtfFontBuilder

        public IRtfFont CreateFont()
        {
            var fontName = FontName;
            if (string.IsNullOrEmpty(fontName))
                fontName = "UnnamedFont_" + FontId;
            return new RtfFont(FontId, FontKind, FontPitch,
                FontCharset, FontCodePage, fontName);
        } // CreateFont

        public void Reset()
        {
            FontIndex = 0;
            FontCharset = 0;
            FontCodePage = 0;
            FontKind = RtfFontKind.Nil;
            FontPitch = RtfFontPitch.Default;
            _fontNameBuffer.Remove(0, _fontNameBuffer.Length);
        } // Reset

        protected override void DoVisitGroup(IRtfGroup group)
        {
            switch (group.Destination)
            {
                case RtfSpec.TagFont:
                case RtfSpec.TagThemeFontLoMajor:
                case RtfSpec.TagThemeFontHiMajor:
                case RtfSpec.TagThemeFontDbMajor:
                case RtfSpec.TagThemeFontBiMajor:
                case RtfSpec.TagThemeFontLoMinor:
                case RtfSpec.TagThemeFontHiMinor:
                case RtfSpec.TagThemeFontDbMinor:
                case RtfSpec.TagThemeFontBiMinor:
                    VisitGroupChildren(group);
                    break;
            }
        } // DoVisitGroup

        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagThemeFontLoMajor:
                case RtfSpec.TagThemeFontHiMajor:
                case RtfSpec.TagThemeFontDbMajor:
                case RtfSpec.TagThemeFontBiMajor:
                case RtfSpec.TagThemeFontLoMinor:
                case RtfSpec.TagThemeFontHiMinor:
                case RtfSpec.TagThemeFontDbMinor:
                case RtfSpec.TagThemeFontBiMinor:
                    // skip and ignore for the moment
                    break;
                case RtfSpec.TagFont:
                    FontId = tag.FullName;
                    FontIndex = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagFontKindNil:
                    FontKind = RtfFontKind.Nil;
                    break;
                case RtfSpec.TagFontKindRoman:
                    FontKind = RtfFontKind.Roman;
                    break;
                case RtfSpec.TagFontKindSwiss:
                    FontKind = RtfFontKind.Swiss;
                    break;
                case RtfSpec.TagFontKindModern:
                    FontKind = RtfFontKind.Modern;
                    break;
                case RtfSpec.TagFontKindScript:
                    FontKind = RtfFontKind.Script;
                    break;
                case RtfSpec.TagFontKindDecor:
                    FontKind = RtfFontKind.Decor;
                    break;
                case RtfSpec.TagFontKindTech:
                    FontKind = RtfFontKind.Tech;
                    break;
                case RtfSpec.TagFontKindBidi:
                    FontKind = RtfFontKind.Bidi;
                    break;
                case RtfSpec.TagFontCharset:
                    FontCharset = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagCodePage:
                    FontCodePage = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagFontPitch:
                    switch (tag.ValueAsNumber)
                    {
                        case 0:
                            FontPitch = RtfFontPitch.Default;
                            break;
                        case 1:
                            FontPitch = RtfFontPitch.Fixed;
                            break;
                        case 2:
                            FontPitch = RtfFontPitch.Variable;
                            break;
                    }
                    break;
            }
        } // DoVisitTag

        protected override void DoVisitText(IRtfText text)
        {
            _fontNameBuffer.Append(text.Text);
        } // DoVisitText
    } // class RtfFontBuilder
}