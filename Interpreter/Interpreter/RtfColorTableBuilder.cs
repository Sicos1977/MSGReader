// -- FILE ------------------------------------------------------------------
// name       : RtfColorTableBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using Itenso.Rtf.Model;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    // ------------------------------------------------------------------------
    public sealed class RtfColorTableBuilder : RtfElementVisitorBase
    {
        // ----------------------------------------------------------------------
        // members
        private readonly RtfColorCollection colorTable;
        private int curBlue;
        private int curGreen;

        private int curRed;

        // ----------------------------------------------------------------------
        public RtfColorTableBuilder(RtfColorCollection colorTable) :
            base(RtfElementVisitorOrder.NonRecursive)
        {
            // we iterate over our children ourselves -> hence non-recursive
            if (colorTable == null)
                throw new ArgumentNullException("colorTable");
            this.colorTable = colorTable;
        } // RtfColorTableBuilder

        // ----------------------------------------------------------------------
        public void Reset()
        {
            colorTable.Clear();
            curRed = 0;
            curGreen = 0;
            curBlue = 0;
        } // Reset

        // ----------------------------------------------------------------------
        protected override void DoVisitGroup(IRtfGroup group)
        {
            if (RtfSpec.TagColorTable.Equals(group.Destination))
                VisitGroupChildren(group);
        } // DoVisitGroup

        // ----------------------------------------------------------------------
        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagColorRed:
                    curRed = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagColorGreen:
                    curGreen = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagColorBlue:
                    curBlue = tag.ValueAsNumber;
                    break;
            }
        } // DoVisitTag

        // ----------------------------------------------------------------------
        protected override void DoVisitText(IRtfText text)
        {
            if (RtfSpec.TagDelimiter.Equals(text.Text))
            {
                colorTable.Add(new RtfColor(curRed, curGreen, curBlue));
                curRed = 0;
                curGreen = 0;
                curBlue = 0;
            }
            else
            {
                throw new RtfColorTableFormatException(Strings.ColorTableUnsupportedText(text.Text));
            }
        } // DoVisitText
    } // class RtfColorBuilder
} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------