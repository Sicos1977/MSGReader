// -- FILE ------------------------------------------------------------------
// name       : RtfText.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfText : RtfElement, IRtfText
	{

		// ----------------------------------------------------------------------
		public RtfText( string text ) :
			base( RtfElementKind.Text )
		{
			if ( text == null )
			{
				throw new ArgumentNullException( "text" );
			}
			this.text = text;
		} // RtfText

		// ----------------------------------------------------------------------
		public string Text
		{
			get { return text; }
		} // Text

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return text;
		} // ToString

		// ----------------------------------------------------------------------
		protected override void DoVisit( IRtfElementVisitor visitor )
		{
			visitor.VisitText( this );
		} // DoVisit

		// ----------------------------------------------------------------------
		protected override bool IsEqual( object obj )
		{
			RtfText compare = obj as RtfText; // guaranteed to be non-null
			return compare != null && base.IsEqual( obj ) &&
				text.Equals( compare.text );
		} // IsEqual

		// ----------------------------------------------------------------------
		protected override int ComputeHashCode()
		{
			return HashTool.AddHashCode( base.ComputeHashCode(), text );
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		// members
		private readonly string text;

	} // class RtfText

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
