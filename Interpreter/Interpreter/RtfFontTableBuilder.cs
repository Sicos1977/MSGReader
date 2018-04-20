// name       : RtfFontTableBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using Itenso.Rtf.Model;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    public sealed class RtfFontTableBuilder : RtfElementVisitorBase
    {
        // members
        private readonly RtfFontBuilder fontBuilder = new RtfFontBuilder();
        private readonly RtfFontCollection fontTable;

        public bool IgnoreDuplicatedFonts { get; } // IgnoreDuplicatedFonts

        public RtfFontTableBuilder(RtfFontCollection fontTable, bool ignoreDuplicatedFonts = true) :
            base(RtfElementVisitorOrder.NonRecursive)
        {
            // we iterate over our children ourselves -> hence non-recursive
            if (fontTable == null)
                throw new ArgumentNullException("fontTable");

            this.fontTable = fontTable;
            IgnoreDuplicatedFonts = ignoreDuplicatedFonts;
        } // RtfFontTableBuilder

        public void Reset()
        {
            fontTable.Clear();
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
                    BuildFontFromGroup(group);
                    break;
                case RtfSpec.TagFontTable:
                    if (group.Contents.Count > 1)
                        if (group.Contents[1].Kind == RtfElementKind.Group)
                        {
                            // the 'new' style where each font resides in a group of its own
                            VisitGroupChildren(group);
                        }
                        else
                        {
                            // the 'old' style where individual fonts are 'terminated' by their
                            // respective name content text (which ends with ';')
                            // -> need to manually iterate from here
                            var childCount = group.Contents.Count;
                            fontBuilder.Reset();
                            for (var i = 1; i < childCount; i++) // skip over the initial \fonttbl tag
                            {
                                group.Contents[i].Visit(fontBuilder);
                                if (fontBuilder.FontName != null)
                                {
                                    // fonts are 'terminated' by their name (as content text)
                                    AddCurrentFont();
                                    fontBuilder.Reset();
                                }
                            }
                            //BuildFontFromGroup( group ); // a single font info
                        }
                    break;
            }
        } // DoVisitGroup

        private void BuildFontFromGroup(IRtfGroup group)
        {
            fontBuilder.Reset();
            fontBuilder.VisitGroup(group);
            AddCurrentFont();
        } // BuildFontFromGroup

        private void AddCurrentFont()
        {
            if (!fontTable.ContainsFontWithId(fontBuilder.FontId))
                fontTable.Add(fontBuilder.CreateFont());
            else if (!IgnoreDuplicatedFonts)
                throw new RtfFontTableFormatException(Strings.DuplicateFont(fontBuilder.FontId));
        } // AddCurrentFont
    } // class RtfFontBuilder
}