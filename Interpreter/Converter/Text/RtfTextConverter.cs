// -- FILE ------------------------------------------------------------------
// name       : RtfTextConverter.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Text;
using System.Globalization;
using Itenso.Rtf.Interpreter;

namespace Itenso.Rtf.Converter.Text
{

	// ------------------------------------------------------------------------
	public class RtfTextConverter : RtfInterpreterListenerBase
	{

		// ----------------------------------------------------------------------
		public const string DefaultTextFileExtension = ".txt";

		// ----------------------------------------------------------------------
		public RtfTextConverter() :
			this( new RtfTextConvertSettings() )
		{
		} // RtfTextConverter

		// ----------------------------------------------------------------------
		public RtfTextConverter( RtfTextConvertSettings settings )
		{
			if ( settings == null )
			{
				throw new ArgumentNullException( "settings" );
			}

			this.settings = settings;
		} // RtfTextConverter

		// ----------------------------------------------------------------------
		public string PlainText
		{
			get { return plainText.ToString(); }
		} // PlainText

		// ----------------------------------------------------------------------
		public RtfTextConvertSettings Settings
		{
			get { return settings; }
		} // Settings

		// ----------------------------------------------------------------------
		public void Clear()
		{
			plainText.Remove( 0, plainText.Length );
		} // Clear

		// ----------------------------------------------------------------------
		protected override void DoBeginDocument( IRtfInterpreterContext context )
		{
			Clear();
		} // DoBeginDocument

		// ----------------------------------------------------------------------
		protected override void DoInsertText( IRtfInterpreterContext context, string text )
		{
			if ( context.CurrentTextFormat == null )
			{
				return;
			}
			if ( !context.CurrentTextFormat.IsHidden || settings.IsShowHiddenText )
			{
				plainText.Append( text );
			}
		} // DoInsertText

		// ----------------------------------------------------------------------
		protected override void DoInsertSpecialChar( IRtfInterpreterContext context, RtfVisualSpecialCharKind kind )
		{
			switch ( kind )
			{
				case RtfVisualSpecialCharKind.Tabulator:
					plainText.Append( settings.TabulatorText );
					break;
				case RtfVisualSpecialCharKind.NonBreakingSpace:
					plainText.Append( settings.NonBreakingSpaceText );
					break;
				case RtfVisualSpecialCharKind.EmSpace:
					plainText.Append( settings.EmSpaceText );
					break;
				case RtfVisualSpecialCharKind.EnSpace:
					plainText.Append( settings.EnSpaceText );
					break;
				case RtfVisualSpecialCharKind.QmSpace:
					plainText.Append( settings.QmSpaceText );
					break;
				case RtfVisualSpecialCharKind.EmDash:
					plainText.Append( settings.EmDashText );
					break;
				case RtfVisualSpecialCharKind.EnDash:
					plainText.Append( settings.EnDashText );
					break;
				case RtfVisualSpecialCharKind.OptionalHyphen:
					plainText.Append( settings.OptionalHyphenText );
					break;
				case RtfVisualSpecialCharKind.NonBreakingHyphen:
					plainText.Append( settings.NonBreakingHyphenText );
					break;
				case RtfVisualSpecialCharKind.Bullet:
					plainText.Append( settings.BulletText );
					break;
				case RtfVisualSpecialCharKind.LeftSingleQuote:
					plainText.Append( settings.LeftSingleQuoteText );
					break;
				case RtfVisualSpecialCharKind.RightSingleQuote:
					plainText.Append( settings.RightSingleQuoteText );
					break;
				case RtfVisualSpecialCharKind.LeftDoubleQuote:
					plainText.Append( settings.LeftDoubleQuoteText );
					break;
				case RtfVisualSpecialCharKind.RightDoubleQuote:
					plainText.Append( settings.RightDoubleQuoteText );
					break;
				default:
					plainText.Append( settings.UnknownSpecialCharText );
					break;
			}
		} // DoInsertSpecialChar

		// ----------------------------------------------------------------------
		protected override void DoInsertBreak( IRtfInterpreterContext context, RtfVisualBreakKind kind )
		{
			switch ( kind )
			{
				case RtfVisualBreakKind.Line:
					plainText.Append( settings.LineBreakText );
					break;
				case RtfVisualBreakKind.Page:
					plainText.Append( settings.PageBreakText );
					break;
				case RtfVisualBreakKind.Paragraph:
					plainText.Append( settings.ParagraphBreakText );
					break;
				case RtfVisualBreakKind.Section:
					plainText.Append( settings.SectionBreakText );
					break;
				default:
					plainText.Append( settings.UnknownBreakText );
					break;
			}
		} // DoInsertBreak

		// ----------------------------------------------------------------------
		protected override void DoInsertImage( IRtfInterpreterContext context,
			RtfVisualImageFormat format,
			int width, int height, int desiredWidth, int desiredHeight,
			int scaleWidthPercent, int scaleHeightPercent,
			string imageDataHex
		)
		{
			string imageFormatText = settings.ImageFormatText;
			if ( string.IsNullOrEmpty( imageFormatText ) )
			{
				return;
			}

			string imageText = string.Format(
				CultureInfo.InvariantCulture,
				imageFormatText,
				format,
				width,
				height,
				desiredWidth,
				desiredHeight,
				scaleWidthPercent,
				scaleHeightPercent,
				imageDataHex );

			plainText.Append( imageText );
		} // DoInsertImage

		// ----------------------------------------------------------------------
		// members
		private readonly StringBuilder plainText = new StringBuilder();
		private readonly RtfTextConvertSettings settings;

	} // class RtfTextConverter

} // namespace Itenso.Rtf.Converter.Text
// -- EOF -------------------------------------------------------------------
