// -- FILE ------------------------------------------------------------------
// name       : RtfHtmlStyle.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.02
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using Itenso.Sys;

namespace Itenso.Rtf.Converter.Html
{

	// ------------------------------------------------------------------------
	public class RtfHtmlStyle : IRtfHtmlStyle
	{

		// ----------------------------------------------------------------------
		public static RtfHtmlStyle Empty = new RtfHtmlStyle();

		// ----------------------------------------------------------------------
		public string ForegroundColor
		{
			get { return foregroundColor; }
			set { foregroundColor = value; }
		} // ForegroundColor

		// ----------------------------------------------------------------------
		public string BackgroundColor
		{
			get { return backgroundColor; }
			set { backgroundColor = value; }
		} // BackgroundColor

		// ----------------------------------------------------------------------
		public string FontFamily
		{
			get { return fontFamily; }
			set { fontFamily = value; }
		} // FontFamily

		// ----------------------------------------------------------------------
		public string FontSize
		{
			get { return fontSize; }
			set { fontSize = value; }
		} // FontSize

		// ----------------------------------------------------------------------
		public bool IsEmpty
		{
			get { return Equals( Empty ); }
		} // IsEmpty

		// ----------------------------------------------------------------------
		public sealed override bool Equals( object obj )
		{
			if ( obj == this )
			{
				return true;
			}
			
			if ( obj == null || GetType() != obj.GetType() )
			{
				return false;
			}

			return IsEqual( obj );
		} // Equals

		// ----------------------------------------------------------------------
		public sealed override int GetHashCode()
		{
			return HashTool.AddHashCode( GetType().GetHashCode(), ComputeHashCode() );
		} // GetHashCode

		// ----------------------------------------------------------------------
		private bool IsEqual( object obj )
		{
			RtfHtmlStyle compare = obj as RtfHtmlStyle; // guaranteed to be non-null
			return
				compare != null &&
				string.Equals( foregroundColor, compare.foregroundColor ) &&
				string.Equals( backgroundColor, compare.backgroundColor ) &&
				string.Equals( fontFamily, compare.fontFamily ) &&
				string.Equals( fontSize, compare.fontSize );
		} // IsEqual

		// ----------------------------------------------------------------------
		private int ComputeHashCode()
		{
			int hash = foregroundColor.GetHashCode();
			hash = HashTool.AddHashCode( hash, backgroundColor );
			hash = HashTool.AddHashCode( hash, fontFamily );
			hash = HashTool.AddHashCode( hash, fontSize );
			return hash;
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		// members
		private string foregroundColor;
		private string backgroundColor;
		private string fontFamily;
		private string fontSize;

	} // class RtfHtmlStyle

} // namespace Itenso.Rtf.Converter.Html
// -- EOF -------------------------------------------------------------------
