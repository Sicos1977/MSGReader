// -- FILE ------------------------------------------------------------------
// name       : RtfHtmlStyleConverter.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.07.13
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Drawing;

namespace Itenso.Rtf.Converter.Html
{

	// ------------------------------------------------------------------------
	public class RtfHtmlStyleConverter : IRtfHtmlStyleConverter
	{

		// ----------------------------------------------------------------------
		public virtual IRtfHtmlStyle TextToHtml( IRtfVisualText visualText )
		{
			if ( visualText == null )
			{
				throw new ArgumentNullException( "visualText" );
			}

			RtfHtmlStyle htmlStyle = new RtfHtmlStyle();

			IRtfTextFormat textFormat = visualText.Format;

			// background color
			Color backgroundColor = textFormat.BackgroundColor.AsDrawingColor;
			if ( backgroundColor.R != 0 || backgroundColor.G != 0 || backgroundColor.B != 0 )
			{
				htmlStyle.BackgroundColor = ColorTranslator.ToHtml( backgroundColor );
			}

			// foreground color
			Color foregroundColor = textFormat.ForegroundColor.AsDrawingColor;
			if ( foregroundColor.R != 0 || foregroundColor.G != 0 || foregroundColor.B != 0 )
			{
				htmlStyle.ForegroundColor = ColorTranslator.ToHtml( foregroundColor );
			}

			// font
			htmlStyle.FontFamily = textFormat.Font.Name;
			if ( textFormat.FontSize > 0 )
			{
				htmlStyle.FontSize = (textFormat.FontSize /2) + "pt";
			}

			return htmlStyle;
		} // TextToHtml

	} // class RtfHtmlStyleConverter

} // namespace Itenso.Rtf.Converter.Html
// -- EOF -------------------------------------------------------------------
