// -- FILE ------------------------------------------------------------------
// name       : RtfTimestampBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{

	// ------------------------------------------------------------------------
	public sealed class RtfTimestampBuilder : RtfElementVisitorBase
	{

		// ----------------------------------------------------------------------
		public RtfTimestampBuilder() :
			base( RtfElementVisitorOrder.BreadthFirst )
		{
			Reset();
		} // RtfTimestampBuilder

		// ----------------------------------------------------------------------
		public void Reset()
		{
			year = 1970;
			month = 1;
			day = 1;
			hour = 0;
			minutes = 0;
			seconds = 0;
		} // Reset

		// ----------------------------------------------------------------------
		public DateTime CreateTimestamp()
		{
			return new DateTime( year, month, day, hour, minutes, seconds );
		} // CreateTimestamp

		// ----------------------------------------------------------------------
		protected override void DoVisitTag( IRtfTag tag )
		{
			switch ( tag.Name )
			{
				case RtfSpec.TagInfoYear:
					year = tag.ValueAsNumber;
					break;
				case RtfSpec.TagInfoMonth:
					month = tag.ValueAsNumber;
					break;
				case RtfSpec.TagInfoDay:
					day = tag.ValueAsNumber;
					break;
				case RtfSpec.TagInfoHour:
					hour = tag.ValueAsNumber;
					break;
				case RtfSpec.TagInfoMinute:
					minutes = tag.ValueAsNumber;
					break;
				case RtfSpec.TagInfoSecond:
					seconds = tag.ValueAsNumber;
					break;
			}
		} // DoVisitTag

		// ----------------------------------------------------------------------
		// members
		private int year;
		private int month;
		private int day;
		private int hour;
		private int minutes;
		private int seconds;

	} // class RtfTimestampBuilder

} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------
