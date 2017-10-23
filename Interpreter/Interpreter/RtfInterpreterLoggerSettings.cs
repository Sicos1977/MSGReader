// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreterLoggerSettings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.05.29
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf.Interpreter
{

	// ------------------------------------------------------------------------
	public class RtfInterpreterLoggerSettings
	{

		// ----------------------------------------------------------------------
		public RtfInterpreterLoggerSettings() :
			this( true )
		{
		} // RtfInterpreterLoggerSettings

		// ----------------------------------------------------------------------
		public RtfInterpreterLoggerSettings( bool enabled )
		{
			Enabled = enabled;
		} // RtfInterpreterLoggerSettings

		// ----------------------------------------------------------------------
		public bool Enabled { get; set; }

		// ----------------------------------------------------------------------
		public string BeginDocumentText
		{
			get { return beginDocumentText; }
			set { beginDocumentText = value; }
		} // BeginDocumentText

		// ----------------------------------------------------------------------
		public string EndDocumentText
		{
			get { return endDocumentText; }
			set { endDocumentText = value; }
		} // EndDocumentText

		// ----------------------------------------------------------------------
		public string TextFormatText
		{
			get { return textFormatText; }
			set { textFormatText = value; }
		} // TextFormatText

		// ----------------------------------------------------------------------
		public string TextOverflowText
		{
			get { return textOverflowText; }
			set { textOverflowText = value; }
		} // TextOverflowText

		// ----------------------------------------------------------------------
		public string SpecialCharFormatText
		{
			get { return specialCharFormatText; }
			set { specialCharFormatText = value; }
		} // SpecialCharFormatText

		// ----------------------------------------------------------------------
		public string BreakFormatText
		{
			get { return breakFormatText; }
			set { breakFormatText = value; }
		} // BreakFormatText

		// ----------------------------------------------------------------------
		public string ImageFormatText
		{
			get { return imageFormatText; }
			set { imageFormatText = value; }
		} // ImageFormatText

		// ----------------------------------------------------------------------
		public int TextMaxLength
		{
			get { return textMaxLength; }
			set { textMaxLength = value; }
		} // TextMaxLength

		// ----------------------------------------------------------------------
		// members
		private string beginDocumentText = Strings.LogBeginDocument;
		private string endDocumentText = Strings.LogEndDocument;
		private string textFormatText = Strings.LogInsertText;
		private string textOverflowText = Strings.LogOverflowText;
		private string specialCharFormatText = Strings.LogInsertChar;
		private string breakFormatText = Strings.LogInsertBreak;
		private string imageFormatText = Strings.LogInsertImage;

		private int textMaxLength = 80;

	} // class RtfInterpreterLoggerSettings

} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------
