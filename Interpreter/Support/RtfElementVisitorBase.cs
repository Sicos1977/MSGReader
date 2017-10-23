﻿// -- FILE ------------------------------------------------------------------
// name       : RtfElementVisitorBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf.Support
{

	// ------------------------------------------------------------------------
	public class RtfElementVisitorBase : IRtfElementVisitor
	{

		// ----------------------------------------------------------------------
		public RtfElementVisitorBase( RtfElementVisitorOrder order )
		{
			this.order = order;
		} // RtfElementVisitorBase

		// ----------------------------------------------------------------------
		public void VisitTag( IRtfTag tag )
		{
			if ( tag != null )
			{
				DoVisitTag( tag );
			}
		} // VisitTag

		// ----------------------------------------------------------------------
		protected virtual void DoVisitTag( IRtfTag tag )
		{
		} // DoVisitTag

		// ----------------------------------------------------------------------
		public void VisitGroup( IRtfGroup group )
		{
			if ( group != null )
			{
				if ( order == RtfElementVisitorOrder.DepthFirst )
				{
					VisitGroupChildren( group );
				}
				DoVisitGroup( group );
				if ( order == RtfElementVisitorOrder.BreadthFirst )
				{
					VisitGroupChildren( group );
				}
			}
		} // VisitGroup

		// ----------------------------------------------------------------------
		protected virtual void DoVisitGroup( IRtfGroup group )
		{
		} // DoVisitGroup

		// ----------------------------------------------------------------------
		protected void VisitGroupChildren( IRtfGroup group )
		{
			foreach ( IRtfElement child in group.Contents )
			{
				child.Visit( this );
			}
		} // VisitGroupChildren

		// ----------------------------------------------------------------------
		public void VisitText( IRtfText text )
		{
			if ( text != null )
			{
				DoVisitText( text );
			}
		} // VisitText

		// ----------------------------------------------------------------------
		protected virtual void DoVisitText( IRtfText text )
		{
		} // DoVisitText

		// ----------------------------------------------------------------------
		// members
		private readonly RtfElementVisitorOrder order;

	} // class RtfElementVisitorBase

} // namespace Itenso.Rtf.Support
// -- EOF -------------------------------------------------------------------
