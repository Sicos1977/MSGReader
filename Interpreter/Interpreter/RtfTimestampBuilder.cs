// name       : RtfTimestampBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    public sealed class RtfTimestampBuilder : RtfElementVisitorBase
    {
        private int _day;
        private int _hour;
        private int _minutes;
        private int _month;
        private int _seconds;

        // Members
        private int _year;

        public RtfTimestampBuilder() :
            base(RtfElementVisitorOrder.BreadthFirst)
        {
            Reset();
        } // RtfTimestampBuilder

        public void Reset()
        {
            _year = 1970;
            _month = 1;
            _day = 1;
            _hour = 0;
            _minutes = 0;
            _seconds = 0;
        } // Reset

        public DateTime CreateTimestamp()
        {
            return new DateTime(_year, _month, _day, _hour, _minutes, _seconds);
        } // CreateTimestamp

        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagInfoYear:
                    _year = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoMonth:
                    _month = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoDay:
                    _day = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoHour:
                    _hour = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoMinute:
                    _minutes = tag.ValueAsNumber;
                    break;
                case RtfSpec.TagInfoSecond:
                    _seconds = tag.ValueAsNumber;
                    break;
            }
        } // DoVisitTag
    } // class RtfTimestampBuilder
}