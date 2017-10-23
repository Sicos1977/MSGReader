// -- FILE ------------------------------------------------------------------
// name       : RtfVisualText.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.22
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfVisualText : RtfVisual, IRtfVisualText
	{

		// ----------------------------------------------------------------------
		public RtfVisualText( string text, IRtfTextFormat format ) :
			base( RtfVisualKind.Text )
		{
			if ( text == null )
			{
				throw new ArgumentNullException( "text" );
			}
			if ( format == null )
			{
				throw new ArgumentNullException( "format" );
			}
			this.text = text;
			this.format = format;
		} // RtfVisualText

		// ----------------------------------------------------------------------
		protected override void DoVisit( IRtfVisualVisitor visitor )
		{
			visitor.VisitText( this );
		} // DoVisit

		// ----------------------------------------------------------------------
		public string Text
		{
			get { return text; }
		} // Text

		// ----------------------------------------------------------------------
		public IRtfTextFormat Format
		{
			get { return format; }
			set
			{
				if ( format == null )
				{
					throw new ArgumentNullException( "value" );
				}
				format = value;
			}
		} // Format

		// ----------------------------------------------------------------------
		protected override bool IsEqual( object obj )
		{
			RtfVisualText compare = obj as RtfVisualText; // guaranteed to be non-null
			return 
				compare != null &&
				base.IsEqual( compare ) &&
				text.Equals( compare.text ) &&
				format.Equals( compare.format );
		} // IsEqual

		// ----------------------------------------------------------------------
		protected override int ComputeHashCode()
		{
			int hash = base.ComputeHashCode();
			hash = HashTool.AddHashCode( hash, text );
			hash = HashTool.AddHashCode( hash, format );
			return hash;
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return "'" + text + "'";
		} // ToString

		// ----------------------------------------------------------------------
		// members
		private readonly string text;
		private IRtfTextFormat format;

	} // class RtfVisualText

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
