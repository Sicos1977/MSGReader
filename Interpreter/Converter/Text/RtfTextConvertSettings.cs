// -- FILE ------------------------------------------------------------------
// name       : RtfTextConvertSettings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.05.29
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;

namespace Itenso.Rtf.Converter.Text
{

	// ------------------------------------------------------------------------
	public class RtfTextConvertSettings
	{

		// ----------------------------------------------------------------------
		public const string SeparatorCr = "\r";
		public const string SeparatorLf = "\n";
		public const string SeparatorCrLf = "\r\n";
		public const string SeparatorLfCr = "\n\r";

		// ----------------------------------------------------------------------
		public RtfTextConvertSettings() :
			this( SeparatorCrLf )
		{
		} // RtfTextConvertSettings

		// ----------------------------------------------------------------------
		public RtfTextConvertSettings( string breakText )
		{
			SetBreakText( breakText );
		} // RtfTextConvertSettings

		// ----------------------------------------------------------------------
		public bool IsShowHiddenText { get; set; }

		// ----------------------------------------------------------------------
		public string TabulatorText
		{
			get { return tabulatorText; }
			set { tabulatorText = value; }
		} // TabulatorText

		// ----------------------------------------------------------------------
		public string NonBreakingSpaceText
		{
			get { return nonBreakingSpaceText; }
			set { nonBreakingSpaceText = value; }
		} // NonBreakingSpaceText

		// ----------------------------------------------------------------------
		public string EmSpaceText
		{
			get { return emSpaceText; }
			set { emSpaceText = value; }
		} // EmSpaceText

		// ----------------------------------------------------------------------
		public string EnSpaceText
		{
			get { return enSpaceText; }
			set { enSpaceText = value; }
		} // EnSpaceText

		// ----------------------------------------------------------------------
		public string QmSpaceText
		{
			get { return qmSpaceText; }
			set { qmSpaceText = value; }
		} // QmSpaceText

		// ----------------------------------------------------------------------
		public string EmDashText
		{
			get { return emDashText; }
			set { emDashText = value; }
		} // EmDashText

		// ----------------------------------------------------------------------
		public string EnDashText
		{
			get { return enDashText; }
			set { enDashText = value; }
		} // EnDashText

		// ----------------------------------------------------------------------
		public string OptionalHyphenText
		{
			get { return optionalHyphenText; }
			set { optionalHyphenText = value; }
		} // OptionalHyphenText

		// ----------------------------------------------------------------------
		public string NonBreakingHyphenText
		{
			get { return nonBreakingHyphenText; }
			set { nonBreakingHyphenText = value; }
		} // NonBreakingHyphenText

		// ----------------------------------------------------------------------
		public string BulletText
		{
			get { return bulletText; }
			set { bulletText = value; }
		} // BulletText

		// ----------------------------------------------------------------------
		public string LeftSingleQuoteText
		{
			get { return leftSingleQuoteText; }
			set { leftSingleQuoteText = value; }
		} // LeftSingleQuoteText

		// ----------------------------------------------------------------------
		public string RightSingleQuoteText
		{
			get { return rightSingleQuoteText; }
			set { rightSingleQuoteText = value; }
		} // RightSingleQuoteText

		// ----------------------------------------------------------------------
		public string LeftDoubleQuoteText
		{
			get { return leftDoubleQuoteText; }
			set { leftDoubleQuoteText = value; }
		} // LeftDoubleQuoteText

		// ----------------------------------------------------------------------
		public string RightDoubleQuoteText
		{
			get { return rightDoubleQuoteText; }
			set { rightDoubleQuoteText = value; }
		} // RightDoubleQuoteText

		// ----------------------------------------------------------------------
		public string UnknownSpecialCharText { get; set; }

		// ----------------------------------------------------------------------
		public string LineBreakText
		{
			get { return lineBreakText; }
			set { lineBreakText = value; }
		} // LineBreakText

		// ----------------------------------------------------------------------
		public string PageBreakText
		{
			get { return pageBreakText; }
			set { pageBreakText = value; }
		} // PageBreakText

		// ----------------------------------------------------------------------
		public string ParagraphBreakText
		{
			get { return paragraphBreakText; }
			set { paragraphBreakText = value; }
		} // ParagraphBreakText

		// ----------------------------------------------------------------------
		public string SectionBreakText
		{
			get { return sectionBreakText; }
			set { sectionBreakText = value; }
		} // SectionBreakText

		// ----------------------------------------------------------------------
		public string UnknownBreakText
		{
			get { return unknownBreakText; }
			set { unknownBreakText = value; }
		} // UnknownBreakText

		// ----------------------------------------------------------------------
		public string ImageFormatText
		{
			get { return imageFormatText; }
			set { imageFormatText = value; }
		} // ImageFormatText

		// ----------------------------------------------------------------------
		public void SetBreakText( string breakText )
		{
			if ( breakText == null )
			{
				throw new ArgumentNullException( "breakText" );
			}

			lineBreakText = breakText;
			pageBreakText = breakText + breakText;
			paragraphBreakText = breakText;
			sectionBreakText = breakText + breakText;
			unknownBreakText = breakText;
		} // SetBreakText

		// ----------------------------------------------------------------------
		// members: hidden text

		// members: special chars
		private string tabulatorText = "\t";
		private string nonBreakingSpaceText = " ";
		private string emSpaceText = " ";
		private string enSpaceText = " ";
		private string qmSpaceText = " ";
		private string emDashText = "-";
		private string enDashText = "-";
		private string optionalHyphenText = "-";
		private string nonBreakingHyphenText = "-";
		private string bulletText = "°";
		private string leftSingleQuoteText = "`";
		private string rightSingleQuoteText = "´";
		private string leftDoubleQuoteText = "``";
		private string rightDoubleQuoteText = "´´";

		// members: breaks
		private string lineBreakText;
		private string pageBreakText;
		private string paragraphBreakText;
		private string sectionBreakText;
		private string unknownBreakText;

		// members: image
		private string imageFormatText = Strings.ImageFormatText;

	} // class RtfTextConvertSettings

} // namespace Itenso.Rtf.Converter.Text
// -- EOF -------------------------------------------------------------------
