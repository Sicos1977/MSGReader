// -- FILE ------------------------------------------------------------------
// name       : RtfHtmlConverter.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.02
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Drawing;
using System.Globalization;
using Itenso.Rtf.Support;
using Itenso.Rtf.Converter.Image;

namespace Itenso.Rtf.Converter.Html
{

	// ------------------------------------------------------------------------
	public class RtfHtmlConverter : RtfVisualVisitorBase
	{

		// ----------------------------------------------------------------------
		public const string DefaultHtmlFileExtension = ".html";

		// ----------------------------------------------------------------------
		public RtfHtmlConverter( IRtfDocument rtfDocument ) :
			this( rtfDocument, new RtfHtmlConvertSettings() )
		{
		} // RtfHtmlConverter

		// ----------------------------------------------------------------------
		public RtfHtmlConverter( IRtfDocument rtfDocument, RtfHtmlConvertSettings settings )
		{
			if ( rtfDocument == null )
			{
				throw new ArgumentNullException( "rtfDocument" );
			}
			if ( settings == null )
			{
				throw new ArgumentNullException( "settings" );
			}

			this.rtfDocument = rtfDocument;
			this.settings = settings;
			specialCharacters = new RtfHtmlSpecialCharCollection( settings.SpecialCharsRepresentation );
		} // RtfHtmlConverter

		// ----------------------------------------------------------------------
		public IRtfDocument RtfDocument
		{
			get { return rtfDocument; }
		} // RtfDocument

		// ----------------------------------------------------------------------
		public RtfHtmlConvertSettings Settings
		{
			get { return settings; }
		} // Settings

		// ----------------------------------------------------------------------
		public IRtfHtmlStyleConverter StyleConverter
		{
			get { return styleConverter; }
			set
			{
				if ( value == null )
				{
					throw new ArgumentNullException( "value" );
				}
				styleConverter = value;
			}
		} // StyleConverter

		// ----------------------------------------------------------------------
		public RtfHtmlSpecialCharCollection SpecialCharacters
		{
			get { return specialCharacters; }
		} // SpecialCharacters

		// ----------------------------------------------------------------------
		public RtfConvertedImageInfoCollection DocumentImages
		{
			get { return documentImages; }
		} // DocumentImages

		// ----------------------------------------------------------------------
		protected HtmlTextWriter Writer
		{
			get { return writer; }
		} // Writer

		// ----------------------------------------------------------------------
		protected RtfHtmlElementPath ElementPath
		{
			get { return elementPath; }
		} // ElementPath

		// ----------------------------------------------------------------------
		protected bool IsInParagraph
		{
			get { return IsInElement( HtmlTextWriterTag.P ); }
		} // IsInParagraph

		// ----------------------------------------------------------------------
		protected bool IsInList
		{
			get { return IsInElement( HtmlTextWriterTag.Ul ) || IsInElement( HtmlTextWriterTag.Ol ); }
		} // IsInList

		// ----------------------------------------------------------------------
		protected bool IsInListItem
		{
			get { return IsInElement( HtmlTextWriterTag.Li ); }
		} // IsInListItem

		// ----------------------------------------------------------------------
		protected virtual string Generator
		{
			get { return generatorName; }
		} // Generator

		// ----------------------------------------------------------------------
		public string Convert()
		{
			string html;
			documentImages.Clear();

			using ( StringWriter stringWriter = new StringWriter() )
			{
				using ( writer = new HtmlTextWriter( stringWriter ) )
				{
					RenderDocumentSection();
					RenderHtmlSection();
				}

				html = stringWriter.ToString();
			}

			if ( elementPath.Count != 0 )
			{
				//logger.Error( "unbalanced element structure" );
			}

			return html;
		} // Convert

		// ----------------------------------------------------------------------
		protected bool IsCurrentElement( HtmlTextWriterTag tag )
		{
			return elementPath.IsCurrent( tag );
		} // IsCurrentElement

		// ----------------------------------------------------------------------
		protected bool IsInElement( HtmlTextWriterTag tag )
		{
			return elementPath.Contains( tag );
		} // IsInElement

		#region TagRendering

		// ----------------------------------------------------------------------
		protected void RenderBeginTag( HtmlTextWriterTag tag )
		{
			Writer.RenderBeginTag( tag );
			elementPath.Push( tag );
		} // RenderBeginTag

		// ----------------------------------------------------------------------
		protected void RenderEndTag()
		{
			RenderEndTag( false );
		} // RenderEndTag

		// ----------------------------------------------------------------------
		protected virtual void RenderEndTag( bool lineBreak )
		{
			Writer.RenderEndTag();
			if ( lineBreak )
			{
				Writer.WriteLine();
			}
			elementPath.Pop();
		} // RenderEndTag

		// ----------------------------------------------------------------------
		protected virtual void RenderTitleTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Title );
		} // RenderTitleTag

		// ----------------------------------------------------------------------
		protected virtual void RenderMetaTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Meta );
		} // RenderMetaTag

		// ----------------------------------------------------------------------
		protected virtual void RenderHtmlTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Html );
		} // RenderHtmlTag

		// ----------------------------------------------------------------------
		protected virtual void RenderLinkTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Link );
		} // RenderLinkTag

		// ----------------------------------------------------------------------
		protected virtual void RenderHeadTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Head );
		} // RenderHeadTag

		// ----------------------------------------------------------------------
		protected virtual void RenderBodyTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Body );
		} // RenderBodyTag

		// ----------------------------------------------------------------------
		protected virtual void RenderLineBreak()
		{
			writer.WriteBreak();
			writer.WriteLine();
		} // RenderLineBreak

		// ----------------------------------------------------------------------
		protected virtual void RenderATag()
		{
			RenderBeginTag( HtmlTextWriterTag.A );
		} // RenderATag

		// ----------------------------------------------------------------------
		protected virtual void RenderPTag()
		{
			RenderBeginTag( HtmlTextWriterTag.P );
		} // RenderPTag

		// ----------------------------------------------------------------------
		protected virtual void RenderBTag()
		{
			RenderBeginTag( HtmlTextWriterTag.B );
		} // RenderBTag

		// ----------------------------------------------------------------------
		protected virtual void RenderITag()
		{
			RenderBeginTag( HtmlTextWriterTag.I );
		} // RenderITag

		// ----------------------------------------------------------------------
		protected virtual void RenderUTag()
		{
			RenderBeginTag( HtmlTextWriterTag.U );
		} // RenderUTag

		// ----------------------------------------------------------------------
		protected virtual void RenderSTag()
		{
			RenderBeginTag( HtmlTextWriterTag.S );
		} // RenderSTag

		// ----------------------------------------------------------------------
		protected virtual void RenderSubTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Sub );
		} // RenderSubTag

		// ----------------------------------------------------------------------
		protected virtual void RenderSupTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Sup );
		} // RenderSupTag

		// ----------------------------------------------------------------------
		protected virtual void RenderSpanTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Span );
		} // RenderSpanTag

		// ----------------------------------------------------------------------
		protected virtual void RenderUlTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Ul );
		} // RenderUlTag

		// ----------------------------------------------------------------------
		protected virtual void RenderOlTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Ol );
		} // RenderOlTag

		// ----------------------------------------------------------------------
		protected virtual void RenderLiTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Li );
		} // RenderLiTag

		// ----------------------------------------------------------------------
		protected virtual void RenderImgTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Img );
		} // RenderImgTag

		// ----------------------------------------------------------------------
		protected virtual void RenderStyleTag()
		{
			RenderBeginTag( HtmlTextWriterTag.Style );
		} // RenderStyleTag

		#endregion // TagRendering

		#region HtmlStructure

		// ----------------------------------------------------------------------
		protected virtual void RenderDocumentHeader()
		{
			if ( string.IsNullOrEmpty( settings.DocumentHeader ) )
			{
				return;
			}

			Writer.WriteLine( settings.DocumentHeader );
		} // RenderDocumentHeader

		// ----------------------------------------------------------------------
		protected virtual void RenderMetaContentType()
		{
			Writer.AddAttribute( "http-equiv", "content-type" );

			string content = "text/html";
			if ( !string.IsNullOrEmpty( settings.CharacterSet ) )
			{
				content = string.Concat( content, "; charset=", settings.CharacterSet );
			}
			Writer.AddAttribute( HtmlTextWriterAttribute.Content, content );
			RenderMetaTag();
			RenderEndTag();
		} // RenderMetaContentType

		// ----------------------------------------------------------------------
		protected virtual void RenderMetaGenerator()
		{
			string generator = Generator;
			if ( string.IsNullOrEmpty( generator ) )
			{
				return;
			}

			Writer.WriteLine();
			Writer.AddAttribute( HtmlTextWriterAttribute.Name, "generator" );
			Writer.AddAttribute( HtmlTextWriterAttribute.Content, generator );
			RenderMetaTag();
			RenderEndTag();
		} // RenderMetaGenerator

		// ----------------------------------------------------------------------
		protected virtual void RenderLinkStyleSheets()
		{
			if ( !settings.HasStyleSheetLinks )
			{
				return;
			}

			foreach ( string styleSheetLink in settings.StyleSheetLinks )
			{
				if ( string.IsNullOrEmpty( styleSheetLink ) )
				{
					continue;
				}

				Writer.WriteLine();
				Writer.AddAttribute( HtmlTextWriterAttribute.Href, styleSheetLink );
				Writer.AddAttribute( HtmlTextWriterAttribute.Type, "text/css" );
				Writer.AddAttribute( HtmlTextWriterAttribute.Rel, "stylesheet" );
				RenderLinkTag();
				RenderEndTag();
			}
		} // RenderLinkStyleSheets

		// ----------------------------------------------------------------------
		protected virtual void RenderHeadAttributes()
		{
			RenderMetaContentType();
			RenderMetaGenerator();
			RenderLinkStyleSheets();
		} // RenderHeadAttributes

		// ----------------------------------------------------------------------
		protected virtual void RenderTitle()
		{
			if ( string.IsNullOrEmpty( settings.Title ) )
			{
				return;
			}

			Writer.WriteLine();
			RenderTitleTag();
			Writer.Write( settings.Title );
			RenderEndTag();
		} // RenderTitle

		// ----------------------------------------------------------------------
		protected virtual void RenderStyles()
		{
			if ( !settings.HasStyles )
			{
				return;
			}

			Writer.WriteLine();
			RenderStyleTag();

			bool firstStyle = true;
			foreach ( IRtfHtmlCssStyle cssStyle in settings.Styles )
			{
				if ( cssStyle.Properties.Count == 0 )
				{
					continue;
				}

				if ( !firstStyle )
				{
					Writer.WriteLine();
				}
				Writer.WriteLine( cssStyle.SelectorName );
				Writer.WriteLine( "{" );
				for ( int i = 0; i < cssStyle.Properties.Count; i++ )
				{
					Writer.WriteLine( string.Format(
						CultureInfo.InvariantCulture,
						"  {0}: {1};",
						cssStyle.Properties.Keys[ i ],
						cssStyle.Properties[ i ] ) );
				}
				Writer.Write( "}" );
				firstStyle = false;
			}

			RenderEndTag();
		} // RenderStyles

		// ----------------------------------------------------------------------
		protected virtual void RenderRtfContent()
		{
			foreach ( IRtfVisual visual in rtfDocument.VisualContent )
			{
				visual.Visit( this );
			}
			EnsureClosedList();
		} // RenderRtfContent

		// ----------------------------------------------------------------------
		protected virtual void BeginParagraph()
		{
			if ( IsInParagraph )
			{
				return;
			}
			RenderPTag();
		} // BeginParagraph

		// ----------------------------------------------------------------------
		protected virtual void EndParagraph()
		{
			if ( !IsInParagraph )
			{
				return;
			}
			RenderEndTag( true );
		} // EndParagraph

		// ----------------------------------------------------------------------
		protected virtual bool OnEnterVisual( IRtfVisual rtfVisual )
		{
			return true;
		} // OnEnterVisual

		// ----------------------------------------------------------------------
		protected virtual void OnLeaveVisual( IRtfVisual rtfVisual )
		{
		} // OnLeaveVisual
		#endregion // HtmlStructure

		#region HtmlFormat

		// ----------------------------------------------------------------------
		protected virtual IRtfHtmlStyle GetHtmlStyle( IRtfVisual rtfVisual )
		{
			IRtfHtmlStyle htmlStyle = RtfHtmlStyle.Empty;

			switch ( rtfVisual.Kind )
			{
				case RtfVisualKind.Text:
					htmlStyle = styleConverter.TextToHtml( rtfVisual as IRtfVisualText );
					break;
			}

			return htmlStyle;
		} // GetHtmlStyle

		// ----------------------------------------------------------------------
		protected virtual string FormatHtmlText( string rtfText )
		{
			string htmlText = HttpUtility.HtmlEncode( rtfText );

			// replace all spaces to non-breaking spaces
			if ( settings.UseNonBreakingSpaces )
			{
				htmlText = htmlText.Replace( " ", nonBreakingSpace );
			}

			return htmlText;
		} // FormatHtmlText

		#endregion // HtmlFormat

		#region RtfVisuals

		// ----------------------------------------------------------------------
		protected override void DoVisitText( IRtfVisualText visualText )
		{
			if ( !EnterVisual( visualText ) )
			{
				return;
			}

			// suppress hidden text
			if ( visualText.Format.IsHidden && settings.IsShowHiddenText == false )
			{
				return;
			}

			IRtfTextFormat textFormat = visualText.Format;
			switch ( textFormat.Alignment )
			{
				case RtfTextAlignment.Left:
					//Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "left" );
					break;
				case RtfTextAlignment.Center:
					Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "center" );
					break;
				case RtfTextAlignment.Right:
					Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "right" );
					break;
				case RtfTextAlignment.Justify:
					Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "justify" );
					break;
			}

			if ( !IsInListItem )
			{
				BeginParagraph();
			}

			// format tags
			if ( textFormat.IsBold )
			{
				RenderBTag();
			}
			if ( textFormat.IsItalic )
			{
				RenderITag();
			}
			if ( textFormat.IsUnderline )
			{
				RenderUTag();
			}
			if ( textFormat.IsStrikeThrough )
			{
				RenderSTag();
			}

			// span with style
			IRtfHtmlStyle htmlStyle = GetHtmlStyle( visualText );
			if ( !htmlStyle.IsEmpty )
			{
				if ( !string.IsNullOrEmpty( htmlStyle.ForegroundColor ) )
				{
					Writer.AddStyleAttribute( HtmlTextWriterStyle.Color, htmlStyle.ForegroundColor );
				}
				if ( !string.IsNullOrEmpty( htmlStyle.BackgroundColor ) )
				{
					Writer.AddStyleAttribute( HtmlTextWriterStyle.BackgroundColor, htmlStyle.BackgroundColor );
				}
				if ( !string.IsNullOrEmpty( htmlStyle.FontFamily ) )
				{
					Writer.AddStyleAttribute( HtmlTextWriterStyle.FontFamily, htmlStyle.FontFamily );
				}
				if ( !string.IsNullOrEmpty( htmlStyle.FontSize ) )
				{
					Writer.AddStyleAttribute( HtmlTextWriterStyle.FontSize, htmlStyle.FontSize );
				}

				RenderSpanTag();
			}

			// visual hyperlink
			bool isHyperlink = false;
			if ( settings.ConvertVisualHyperlinks )
			{
				string href = ConvertVisualHyperlink( visualText.Text );
				if ( !string.IsNullOrEmpty( href ) )
				{
					isHyperlink = true;
					Writer.AddAttribute( HtmlTextWriterAttribute.Href, href );
					RenderATag();
				}
			}

			// subscript and superscript
			if ( textFormat.SuperScript < 0 )
			{
				RenderSubTag();
			}
			else if ( textFormat.SuperScript > 0 )
			{
				RenderSupTag();
			}

			string htmlText = FormatHtmlText( visualText.Text );
			Writer.Write( htmlText );

			// subscript and superscript
			if ( textFormat.SuperScript < 0 )
			{
				RenderEndTag(); // sub
			}
			else if ( textFormat.SuperScript > 0 )
			{
				RenderEndTag(); // sup
			}

			// visual hyperlink
			if ( isHyperlink )
			{
				RenderEndTag(); // a
			}

			// span with style
			if ( !htmlStyle.IsEmpty )
			{
				RenderEndTag();
			}

			// format tags
			if ( textFormat.IsStrikeThrough )
			{
				RenderEndTag(); // s
			}
			if ( textFormat.IsUnderline )
			{
				RenderEndTag(); // u
			}
			if ( textFormat.IsItalic )
			{
				RenderEndTag(); // i
			}
			if ( textFormat.IsBold )
			{
				RenderEndTag(); // b
			}

			LeaveVisual( visualText );
		} // DoVisitText

		// ----------------------------------------------------------------------
		protected override void DoVisitImage( IRtfVisualImage visualImage )
		{
			if ( !EnterVisual( visualImage ) )
			{
				return;
			}

			switch ( visualImage.Alignment )
			{
				case RtfTextAlignment.Left:
					//Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "left" );
					break;
				case RtfTextAlignment.Center:
					Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "center" );
					break;
				case RtfTextAlignment.Right:
					Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "right" );
					break;
				case RtfTextAlignment.Justify:
					Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "justify" );
					break;
			}

			BeginParagraph();

			int imageIndex = documentImages.Count + 1;
			string fileName = settings.GetImageUrl( imageIndex, visualImage.Format );
			int width = settings.ImageAdapter.CalcImageWidth( visualImage.Format, visualImage.Width,
				visualImage.DesiredWidth, visualImage.ScaleWidthPercent );
			int height = settings.ImageAdapter.CalcImageHeight( visualImage.Format, visualImage.Height,
				visualImage.DesiredHeight, visualImage.ScaleHeightPercent );

			Writer.AddAttribute( HtmlTextWriterAttribute.Width, width.ToString( CultureInfo.InvariantCulture ) );
			Writer.AddAttribute( HtmlTextWriterAttribute.Height, height.ToString( CultureInfo.InvariantCulture ) );
			string htmlFileName = HttpUtility.HtmlEncode( fileName );
			Writer.AddAttribute( HtmlTextWriterAttribute.Src, htmlFileName, false );
			RenderImgTag();
			RenderEndTag();

			documentImages.Add( new RtfConvertedImageInfo(
				htmlFileName,
				settings.ImageAdapter.TargetFormat,
				new Size( width, height ) ) );

			LeaveVisual( visualImage );
		} // DoVisitImage

		// ----------------------------------------------------------------------
		protected override void DoVisitSpecial( IRtfVisualSpecialChar visualSpecialChar )
		{
			if ( !EnterVisual( visualSpecialChar ) )
			{
				return;
			}

			switch ( visualSpecialChar.CharKind )
			{
				case RtfVisualSpecialCharKind.ParagraphNumberBegin:
					isInParagraphNumber = true;
					break;
				case RtfVisualSpecialCharKind.ParagraphNumberEnd:
					isInParagraphNumber = false;
					break;
				default:
					if ( SpecialCharacters.ContainsKey( visualSpecialChar.CharKind ) )
					{
						Writer.Write( SpecialCharacters[ visualSpecialChar.CharKind ] );
					}
					break;
			}

			LeaveVisual( visualSpecialChar );
		} // DoVisitSpecial

		// ----------------------------------------------------------------------
		protected override void DoVisitBreak( IRtfVisualBreak visualBreak )
		{
			if ( !EnterVisual( visualBreak ) )
			{
				return;
			}

			switch ( visualBreak.BreakKind )
			{
				case RtfVisualBreakKind.Line:
					RenderLineBreak();
					break;
				case RtfVisualBreakKind.Page:
					break;
				case RtfVisualBreakKind.Paragraph:
					if ( IsInParagraph )
					{
						EndParagraph(); // close paragraph
					}
					else if ( IsInListItem )
					{
						EndParagraph();
						RenderEndTag( true ); // close list item
					}
					else
					{
						BeginParagraph();
						Writer.Write( nonBreakingSpace );
						EndParagraph();
					}
					break;
				case RtfVisualBreakKind.Section:
					break;
			}

			LeaveVisual( visualBreak );
		} // DoVisitBreak

		#endregion // RtfVisuals

		// ----------------------------------------------------------------------
		private string ConvertVisualHyperlink( string text )
		{
			if ( string.IsNullOrEmpty( text ) )
			{
				return null;
			}

			if ( hyperlinkRegEx == null )
			{
				if ( string.IsNullOrEmpty( settings.VisualHyperlinkPattern ) )
				{
					return null;
				}
				hyperlinkRegEx = new Regex( settings.VisualHyperlinkPattern );
			}

			return hyperlinkRegEx.IsMatch( text ) ? text : null;
		} // ConvertVisualHyperlink

		// ----------------------------------------------------------------------
		private void RenderDocumentSection()
		{
			if ( ( settings.ConvertScope & RtfHtmlConvertScope.Document ) != RtfHtmlConvertScope.Document )
			{
				return;
			}

			RenderDocumentHeader();
		} // RenderDocumentSection

		// ----------------------------------------------------------------------
		private void RenderHtmlSection()
		{
			if ( ( settings.ConvertScope & RtfHtmlConvertScope.Html ) == RtfHtmlConvertScope.Html )
			{
				RenderHtmlTag();
			}

			RenderHeadSection();
			RenderBodySection();

			if ( ( settings.ConvertScope & RtfHtmlConvertScope.Html ) == RtfHtmlConvertScope.Html )
			{
				RenderEndTag( true );
			}
		} // RenderHtmlSection

		// ----------------------------------------------------------------------
		private void RenderHeadSection()
		{
			if ( ( settings.ConvertScope & RtfHtmlConvertScope.Head ) != RtfHtmlConvertScope.Head )
			{
				return;
			}

			RenderHeadTag();
			RenderHeadAttributes();
			RenderTitle();
			RenderStyles();
			RenderEndTag( true );
		} // RenderHeadSection

		// ----------------------------------------------------------------------
		private void RenderBodySection()
		{
			if ( ( settings.ConvertScope & RtfHtmlConvertScope.Body ) == RtfHtmlConvertScope.Body )
			{
				RenderBodyTag();
			}

			if ( ( settings.ConvertScope & RtfHtmlConvertScope.Content ) == RtfHtmlConvertScope.Content )
			{
				RenderRtfContent();
			}

			if ( ( settings.ConvertScope & RtfHtmlConvertScope.Body ) == RtfHtmlConvertScope.Body )
			{
				RenderEndTag();
			}
		} // RenderBodySection

		// ----------------------------------------------------------------------
		private bool EnterVisual( IRtfVisual rtfVisual )
		{
			bool openList = EnsureOpenList( rtfVisual );
			if ( openList )
			{
				return false;
			}

			EnsureClosedList( rtfVisual );
			return OnEnterVisual( rtfVisual );
		} // EnterVisual

		// ----------------------------------------------------------------------
		private void LeaveVisual( IRtfVisual rtfVisual )
		{
			OnLeaveVisual( rtfVisual );
			lastVisual = rtfVisual;
		} // LeaveVisual

		// ----------------------------------------------------------------------
		// returns true if visual is in list
		private bool EnsureOpenList( IRtfVisual rtfVisual )
		{
			IRtfVisualText visualText = rtfVisual as IRtfVisualText;
			if ( visualText == null || !isInParagraphNumber )
			{
				return false;
			}

			if ( !IsInList )
			{
				bool unsortedList = unsortedListValue.Equals( visualText.Text );
				if ( unsortedList )
				{
					RenderUlTag(); // unsorted list
				}
				else
				{
					RenderOlTag(); // ordered list
				}
			}

			RenderLiTag();

			return true;
		} // EnsureOpenList

		// ----------------------------------------------------------------------
		private void EnsureClosedList()
		{
			if ( lastVisual == null )
			{
				return;
			}
			EnsureClosedList( lastVisual );
		} // EnsureClosedList

		// ----------------------------------------------------------------------
		private void EnsureClosedList( IRtfVisual rtfVisual )
		{
			if ( !IsInList )
			{
				return; // not in list
			}

			IRtfVisualBreak previousParagraph = lastVisual as IRtfVisualBreak;
			if ( previousParagraph == null || previousParagraph.BreakKind != RtfVisualBreakKind.Paragraph )
			{
				return; // is not following to a pragraph
			}

			IRtfVisualSpecialChar specialChar = rtfVisual as IRtfVisualSpecialChar;
			if ( specialChar == null ||
				specialChar.CharKind != RtfVisualSpecialCharKind.ParagraphNumberBegin )
			{
				RenderEndTag( true ); // close ul/ol list
			}
		} // EnsureClosedList

		// ----------------------------------------------------------------------
		// members
		private readonly RtfConvertedImageInfoCollection documentImages = new RtfConvertedImageInfoCollection();
		private readonly RtfHtmlElementPath elementPath = new RtfHtmlElementPath();
		private readonly IRtfDocument rtfDocument;
		private readonly RtfHtmlConvertSettings settings;
		private IRtfHtmlStyleConverter styleConverter = new RtfHtmlStyleConverter();
		private readonly RtfHtmlSpecialCharCollection specialCharacters;
		private HtmlTextWriter writer;
		private IRtfVisual lastVisual;
		private bool isInParagraphNumber;

		private const string generatorName = "Rtf2Html Converter";
		private const string nonBreakingSpace = "&nbsp;";
		private const string unsortedListValue = "·";

		private Regex hyperlinkRegEx;

	} // class RtfHtmlConverter

} // namespace Itenso.Rtf.Converter.Html
// -- EOF -------------------------------------------------------------------
