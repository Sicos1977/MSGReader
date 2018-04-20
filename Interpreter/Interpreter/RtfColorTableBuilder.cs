// name       : RtfColorTableBuilder.cs
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
    public sealed class RtfColorTableBuilder : RtfElementVisitorBase
    {
        // Members
        private readonly RtfColorCollection _colorTable;
        private int _curBlue;
        private int _curGreen;

        private int _curRed;

        public RtfColorTableBuilder(RtfColorCollection colorTable) :
            base(RtfElementVisitorOrder.NonRecursive)
        {
            // we iterate over our children ourselves -> hence non-recursive
            if (colorTable == null)
                throw new ArgumentNullException(nameof(colorTable));
            _colorTable = colorTable;
        } // RtfColorTableBuilder

        public void Reset()
        {
            _colorTable.Clear();
            _curRed = 0;
            _curGreen = 0;
            _curBlue = 0;
        } // Reset

        protected override void DoVisitGroup(IRtfGroup group)
        {
            if (RtfSpec.TagColorTable.Equals(group.Destination))
                VisitGroupChildren(group);
        } // DoVisitGroup

        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagColorRed:
                    _curRed = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagColorGreen:
                    _curGreen = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagColorBlue:
                    _curBlue = tag.ValueAsNumber;
                    break;
            }
        } // DoVisitTag

        protected override void DoVisitText(IRtfText text)
        {
            if (RtfSpec.TagDelimiter.Equals(text.Text))
            {
                _colorTable.Add(new RtfColor(_curRed, _curGreen, _curBlue));
                _curRed = 0;
                _curGreen = 0;
                _curBlue = 0;
            }
            else
            {
                throw new RtfColorTableFormatException(Strings.ColorTableUnsupportedText(text.Text));
            }
        } // DoVisitText
    } // class RtfColorBuilder
}