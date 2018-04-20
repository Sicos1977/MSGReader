// name       : RtfHtmlConverter.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.02
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using Itenso.Rtf.Converter.Image;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Converter.Html
{
    public class RtfHtmlConverter : RtfVisualVisitorBase
    {
        public const string DefaultHtmlFileExtension = ".html";

        private const string GeneratorName = "Rtf2Html Converter";
        private const string NonBreakingSpace = "&nbsp;";
        private const string UnsortedListValue = "·";

        // Members

        private Regex _hyperlinkRegEx;
        private bool _isInParagraphNumber;
        private IRtfVisual _lastVisual;
        private IRtfHtmlStyleConverter _styleConverter = new RtfHtmlStyleConverter();

        public IRtfDocument RtfDocument { get; } // RtfDocument

        public RtfHtmlConvertSettings Settings { get; } // Settings

        public IRtfHtmlStyleConverter StyleConverter
        {
            get { return _styleConverter; }
            set
            {
                if (value == null)
                    throw new ArgumentNullException(nameof(value));
                _styleConverter = value;
            }
        } // StyleConverter

        public RtfHtmlSpecialCharCollection SpecialCharacters { get; } // SpecialCharacters

        public RtfConvertedImageInfoCollection DocumentImages { get; } = new RtfConvertedImageInfoCollection();

// DocumentImages

        protected HtmlTextWriter Writer { get; private set; } // Writer

        protected RtfHtmlElementPath ElementPath { get; } = new RtfHtmlElementPath();

// ElementPath

        protected bool IsInParagraph => IsInElement(HtmlTextWriterTag.P);
        // IsInParagraph

        protected bool IsInList => IsInElement(HtmlTextWriterTag.Ul) || IsInElement(HtmlTextWriterTag.Ol);
        // IsInList

        protected bool IsInListItem => IsInElement(HtmlTextWriterTag.Li);
        // IsInListItem

        protected virtual string Generator => GeneratorName;
        // Generator

        public RtfHtmlConverter(IRtfDocument rtfDocument) :
            this(rtfDocument, new RtfHtmlConvertSettings())
        {
        } // RtfHtmlConverter

        public RtfHtmlConverter(IRtfDocument rtfDocument, RtfHtmlConvertSettings settings)
        {
            if (rtfDocument == null)
                throw new ArgumentNullException(nameof(rtfDocument));
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            RtfDocument = rtfDocument;
            Settings = settings;
            SpecialCharacters = new RtfHtmlSpecialCharCollection(settings.SpecialCharsRepresentation);
        } // RtfHtmlConverter

        public string Convert()
        {
            string html;
            DocumentImages.Clear();

            using (var stringWriter = new StringWriter())
            {
                using (Writer = new HtmlTextWriter(stringWriter))
                {
                    RenderDocumentSection();
                    RenderHtmlSection();
                }

                html = stringWriter.ToString();
            }

            if (ElementPath.Count != 0)
            {
                //logger.Error( "unbalanced element structure" );
            }

            return html;
        } // Convert

        protected bool IsCurrentElement(HtmlTextWriterTag tag)
        {
            return ElementPath.IsCurrent(tag);
        } // IsCurrentElement

        protected bool IsInElement(HtmlTextWriterTag tag)
        {
            return ElementPath.Contains(tag);
        } // IsInElement

        private string ConvertVisualHyperlink(string text)
        {
            if (string.IsNullOrEmpty(text))
                return null;

            if (_hyperlinkRegEx == null)
            {
                if (string.IsNullOrEmpty(Settings.VisualHyperlinkPattern))
                    return null;
                _hyperlinkRegEx = new Regex(Settings.VisualHyperlinkPattern);
            }

            return _hyperlinkRegEx.IsMatch(text) ? text : null;
        } // ConvertVisualHyperlink

        private void RenderDocumentSection()
        {
            if ((Settings.ConvertScope & RtfHtmlConvertScope.Document) != RtfHtmlConvertScope.Document)
                return;

            RenderDocumentHeader();
        } // RenderDocumentSection

        private void RenderHtmlSection()
        {
            if ((Settings.ConvertScope & RtfHtmlConvertScope.Html) == RtfHtmlConvertScope.Html)
                RenderHtmlTag();

            RenderHeadSection();
            RenderBodySection();

            if ((Settings.ConvertScope & RtfHtmlConvertScope.Html) == RtfHtmlConvertScope.Html)
                RenderEndTag(true);
        } // RenderHtmlSection

        private void RenderHeadSection()
        {
            if ((Settings.ConvertScope & RtfHtmlConvertScope.Head) != RtfHtmlConvertScope.Head)
                return;

            RenderHeadTag();
            RenderHeadAttributes();
            RenderTitle();
            RenderStyles();
            RenderEndTag(true);
        } // RenderHeadSection

        private void RenderBodySection()
        {
            if ((Settings.ConvertScope & RtfHtmlConvertScope.Body) == RtfHtmlConvertScope.Body)
                RenderBodyTag();

            if ((Settings.ConvertScope & RtfHtmlConvertScope.Content) == RtfHtmlConvertScope.Content)
                RenderRtfContent();

            if ((Settings.ConvertScope & RtfHtmlConvertScope.Body) == RtfHtmlConvertScope.Body)
                RenderEndTag();
        } // RenderBodySection

        private bool EnterVisual(IRtfVisual rtfVisual)
        {
            var openList = EnsureOpenList(rtfVisual);
            if (openList)
                return false;

            EnsureClosedList(rtfVisual);
            return OnEnterVisual(rtfVisual);
        } // EnterVisual

        private void LeaveVisual(IRtfVisual rtfVisual)
        {
            OnLeaveVisual(rtfVisual);
            _lastVisual = rtfVisual;
        } // LeaveVisual

        // returns true if visual is in list
        private bool EnsureOpenList(IRtfVisual rtfVisual)
        {
            var visualText = rtfVisual as IRtfVisualText;
            if (visualText == null || !_isInParagraphNumber)
                return false;

            if (!IsInList)
            {
                var unsortedList = UnsortedListValue.Equals(visualText.Text);
                if (unsortedList)
                    RenderUlTag(); // unsorted list
                else
                    RenderOlTag(); // ordered list
            }

            RenderLiTag();

            return true;
        } // EnsureOpenList

        private void EnsureClosedList()
        {
            if (_lastVisual == null)
                return;
            EnsureClosedList(_lastVisual);
        } // EnsureClosedList

        private void EnsureClosedList(IRtfVisual rtfVisual)
        {
            if (!IsInList)
                return; // not in list

            var previousParagraph = _lastVisual as IRtfVisualBreak;
            if (previousParagraph == null || previousParagraph.BreakKind != RtfVisualBreakKind.Paragraph)
                return; // is not following to a pragraph

            var specialChar = rtfVisual as IRtfVisualSpecialChar;
            if (specialChar == null ||
                specialChar.CharKind != RtfVisualSpecialCharKind.ParagraphNumberBegin)
                RenderEndTag(true); // close ul/ol list
        } // EnsureClosedList

        #region TagRendering
        protected void RenderBeginTag(HtmlTextWriterTag tag)
        {
            Writer.RenderBeginTag(tag);
            ElementPath.Push(tag);
        } // RenderBeginTag

        protected void RenderEndTag()
        {
            RenderEndTag(false);
        } // RenderEndTag

        protected virtual void RenderEndTag(bool lineBreak)
        {
            Writer.RenderEndTag();
            if (lineBreak)
                Writer.WriteLine();
            ElementPath.Pop();
        } // RenderEndTag

        protected virtual void RenderTitleTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Title);
        } // RenderTitleTag

        protected virtual void RenderMetaTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Meta);
        } // RenderMetaTag

        protected virtual void RenderHtmlTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Html);
        } // RenderHtmlTag

        protected virtual void RenderLinkTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Link);
        } // RenderLinkTag

        protected virtual void RenderHeadTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Head);
        } // RenderHeadTag

        protected virtual void RenderBodyTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Body);
        } // RenderBodyTag

        protected virtual void RenderLineBreak()
        {
            Writer.WriteBreak();
            Writer.WriteLine();
        } // RenderLineBreak

        protected virtual void RenderATag()
        {
            RenderBeginTag(HtmlTextWriterTag.A);
        } // RenderATag

        protected virtual void RenderPTag()
        {
            RenderBeginTag(HtmlTextWriterTag.P);
        } // RenderPTag

        protected virtual void RenderBTag()
        {
            RenderBeginTag(HtmlTextWriterTag.B);
        } // RenderBTag

        protected virtual void RenderITag()
        {
            RenderBeginTag(HtmlTextWriterTag.I);
        } // RenderITag

        protected virtual void RenderUTag()
        {
            RenderBeginTag(HtmlTextWriterTag.U);
        } // RenderUTag

        protected virtual void RenderSTag()
        {
            RenderBeginTag(HtmlTextWriterTag.S);
        } // RenderSTag

        protected virtual void RenderSubTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Sub);
        } // RenderSubTag

        protected virtual void RenderSupTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Sup);
        } // RenderSupTag

        protected virtual void RenderSpanTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Span);
        } // RenderSpanTag

        protected virtual void RenderUlTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Ul);
        } // RenderUlTag

        protected virtual void RenderOlTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Ol);
        } // RenderOlTag

        protected virtual void RenderLiTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Li);
        } // RenderLiTag

        protected virtual void RenderImgTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Img);
        } // RenderImgTag

        protected virtual void RenderStyleTag()
        {
            RenderBeginTag(HtmlTextWriterTag.Style);
        } // RenderStyleTag
        #endregion // TagRendering

        #region HtmlStructure
        protected virtual void RenderDocumentHeader()
        {
            if (string.IsNullOrEmpty(Settings.DocumentHeader))
                return;

            Writer.WriteLine(Settings.DocumentHeader);
        } // RenderDocumentHeader

        protected virtual void RenderMetaContentType()
        {
            Writer.AddAttribute("http-equiv", "content-type");

            var content = "text/html";
            if (!string.IsNullOrEmpty(Settings.CharacterSet))
                content = string.Concat(content, "; charset=", Settings.CharacterSet);
            Writer.AddAttribute(HtmlTextWriterAttribute.Content, content);
            RenderMetaTag();
            RenderEndTag();
        } // RenderMetaContentType

        protected virtual void RenderMetaGenerator()
        {
            var generator = Generator;
            if (string.IsNullOrEmpty(generator))
                return;

            Writer.WriteLine();
            Writer.AddAttribute(HtmlTextWriterAttribute.Name, "generator");
            Writer.AddAttribute(HtmlTextWriterAttribute.Content, generator);
            RenderMetaTag();
            RenderEndTag();
        } // RenderMetaGenerator

        protected virtual void RenderLinkStyleSheets()
        {
            if (!Settings.HasStyleSheetLinks)
                return;

            foreach (var styleSheetLink in Settings.StyleSheetLinks)
            {
                if (string.IsNullOrEmpty(styleSheetLink))
                    continue;

                Writer.WriteLine();
                Writer.AddAttribute(HtmlTextWriterAttribute.Href, styleSheetLink);
                Writer.AddAttribute(HtmlTextWriterAttribute.Type, "text/css");
                Writer.AddAttribute(HtmlTextWriterAttribute.Rel, "stylesheet");
                RenderLinkTag();
                RenderEndTag();
            }
        } // RenderLinkStyleSheets

        protected virtual void RenderHeadAttributes()
        {
            RenderMetaContentType();
            RenderMetaGenerator();
            RenderLinkStyleSheets();
        } // RenderHeadAttributes

        protected virtual void RenderTitle()
        {
            if (string.IsNullOrEmpty(Settings.Title))
                return;

            Writer.WriteLine();
            RenderTitleTag();
            Writer.Write(Settings.Title);
            RenderEndTag();
        } // RenderTitle

        protected virtual void RenderStyles()
        {
            if (!Settings.HasStyles)
                return;

            Writer.WriteLine();
            RenderStyleTag();

            var firstStyle = true;
            foreach (IRtfHtmlCssStyle cssStyle in Settings.Styles)
            {
                if (cssStyle.Properties.Count == 0)
                    continue;

                if (!firstStyle)
                    Writer.WriteLine();
                Writer.WriteLine(cssStyle.SelectorName);
                Writer.WriteLine("{");
                for (var i = 0; i < cssStyle.Properties.Count; i++)
                    Writer.WriteLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "  {0}: {1};",
                        cssStyle.Properties.Keys[i],
                        cssStyle.Properties[i]));
                Writer.Write("}");
                firstStyle = false;
            }

            RenderEndTag();
        } // RenderStyles

        protected virtual void RenderRtfContent()
        {
            foreach (IRtfVisual visual in RtfDocument.VisualContent)
                visual.Visit(this);
            EnsureClosedList();
        } // RenderRtfContent

        protected virtual void BeginParagraph()
        {
            if (IsInParagraph)
                return;
            RenderPTag();
        } // BeginParagraph

        protected virtual void EndParagraph()
        {
            if (!IsInParagraph)
                return;
            RenderEndTag(true);
        } // EndParagraph

        protected virtual bool OnEnterVisual(IRtfVisual rtfVisual)
        {
            return true;
        } // OnEnterVisual

        protected virtual void OnLeaveVisual(IRtfVisual rtfVisual)
        {
        } // OnLeaveVisual
        #endregion // HtmlStructure

        #region HtmlFormat
        protected virtual IRtfHtmlStyle GetHtmlStyle(IRtfVisual rtfVisual)
        {
            IRtfHtmlStyle htmlStyle = RtfHtmlStyle.Empty;

            switch (rtfVisual.Kind)
            {
                case RtfVisualKind.Text:
                    htmlStyle = _styleConverter.TextToHtml(rtfVisual as IRtfVisualText);
                    break;
            }

            return htmlStyle;
        } // GetHtmlStyle

        protected virtual string FormatHtmlText(string rtfText)
        {
            var htmlText = HttpUtility.HtmlEncode(rtfText);

            // replace all spaces to non-breaking spaces
            if (Settings.UseNonBreakingSpaces)
                htmlText = htmlText?.Replace(" ", NonBreakingSpace);

            return htmlText;
        } // FormatHtmlText
        #endregion // HtmlFormat

        #region RtfVisuals
        protected override void DoVisitText(IRtfVisualText visualText)
        {
            if (!EnterVisual(visualText))
                return;

            // suppress hidden text
            if (visualText.Format.IsHidden && Settings.IsShowHiddenText == false)
                return;

            var textFormat = visualText.Format;
            switch (textFormat.Alignment)
            {
                case RtfTextAlignment.Left:
                    //Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "left" );
                    break;
                case RtfTextAlignment.Center:
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.TextAlign, "center");
                    break;
                case RtfTextAlignment.Right:
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.TextAlign, "right");
                    break;
                case RtfTextAlignment.Justify:
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.TextAlign, "justify");
                    break;
            }

            if (!IsInListItem)
                BeginParagraph();

            // format tags
            if (textFormat.IsBold)
                RenderBTag();
            if (textFormat.IsItalic)
                RenderITag();
            if (textFormat.IsUnderline)
                RenderUTag();
            if (textFormat.IsStrikeThrough)
                RenderSTag();

            // span with style
            var htmlStyle = GetHtmlStyle(visualText);
            if (!htmlStyle.IsEmpty)
            {
                if (!string.IsNullOrEmpty(htmlStyle.ForegroundColor))
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.Color, htmlStyle.ForegroundColor);
                if (!string.IsNullOrEmpty(htmlStyle.BackgroundColor))
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.BackgroundColor, htmlStyle.BackgroundColor);
                if (!string.IsNullOrEmpty(htmlStyle.FontFamily))
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.FontFamily, htmlStyle.FontFamily);
                if (!string.IsNullOrEmpty(htmlStyle.FontSize))
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.FontSize, htmlStyle.FontSize);

                RenderSpanTag();
            }

            // visual hyperlink
            var isHyperlink = false;
            if (Settings.ConvertVisualHyperlinks)
            {
                var href = ConvertVisualHyperlink(visualText.Text);
                if (!string.IsNullOrEmpty(href))
                {
                    isHyperlink = true;
                    Writer.AddAttribute(HtmlTextWriterAttribute.Href, href);
                    RenderATag();
                }
            }

            // subscript and superscript
            if (textFormat.SuperScript < 0)
                RenderSubTag();
            else if (textFormat.SuperScript > 0)
                RenderSupTag();

            var htmlText = FormatHtmlText(visualText.Text);
            Writer.Write(htmlText);

            // subscript and superscript
            if (textFormat.SuperScript < 0)
                RenderEndTag(); // sub
            else if (textFormat.SuperScript > 0)
                RenderEndTag(); // sup

            // visual hyperlink
            if (isHyperlink)
                RenderEndTag(); // a

            // span with style
            if (!htmlStyle.IsEmpty)
                RenderEndTag();

            // format tags
            if (textFormat.IsStrikeThrough)
                RenderEndTag(); // s
            if (textFormat.IsUnderline)
                RenderEndTag(); // u
            if (textFormat.IsItalic)
                RenderEndTag(); // i
            if (textFormat.IsBold)
                RenderEndTag(); // b

            LeaveVisual(visualText);
        } // DoVisitText

        protected override void DoVisitImage(IRtfVisualImage visualImage)
        {
            if (!EnterVisual(visualImage))
                return;

            switch (visualImage.Alignment)
            {
                case RtfTextAlignment.Left:
                    //Writer.AddStyleAttribute( HtmlTextWriterStyle.TextAlign, "left" );
                    break;
                case RtfTextAlignment.Center:
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.TextAlign, "center");
                    break;
                case RtfTextAlignment.Right:
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.TextAlign, "right");
                    break;
                case RtfTextAlignment.Justify:
                    Writer.AddStyleAttribute(HtmlTextWriterStyle.TextAlign, "justify");
                    break;
            }

            BeginParagraph();

            var imageIndex = DocumentImages.Count + 1;
            var fileName = Settings.GetImageUrl(imageIndex, visualImage.Format);
            var width = Settings.ImageAdapter.CalcImageWidth(visualImage.Format, visualImage.Width,
                visualImage.DesiredWidth, visualImage.ScaleWidthPercent);
            var height = Settings.ImageAdapter.CalcImageHeight(visualImage.Format, visualImage.Height,
                visualImage.DesiredHeight, visualImage.ScaleHeightPercent);

            Writer.AddAttribute(HtmlTextWriterAttribute.Width, width.ToString(CultureInfo.InvariantCulture));
            Writer.AddAttribute(HtmlTextWriterAttribute.Height, height.ToString(CultureInfo.InvariantCulture));
            var htmlFileName = HttpUtility.HtmlEncode(fileName);
            Writer.AddAttribute(HtmlTextWriterAttribute.Src, htmlFileName, false);
            RenderImgTag();
            RenderEndTag();

            DocumentImages.Add(new RtfConvertedImageInfo(
                htmlFileName,
                Settings.ImageAdapter.TargetFormat,
                new Size(width, height)));

            LeaveVisual(visualImage);
        } // DoVisitImage

        protected override void DoVisitSpecial(IRtfVisualSpecialChar visualSpecialChar)
        {
            if (!EnterVisual(visualSpecialChar))
                return;

            switch (visualSpecialChar.CharKind)
            {
                case RtfVisualSpecialCharKind.ParagraphNumberBegin:
                    _isInParagraphNumber = true;
                    break;
                case RtfVisualSpecialCharKind.ParagraphNumberEnd:
                    _isInParagraphNumber = false;
                    break;
                default:
                    if (SpecialCharacters.ContainsKey(visualSpecialChar.CharKind))
                        Writer.Write(SpecialCharacters[visualSpecialChar.CharKind]);
                    break;
            }

            LeaveVisual(visualSpecialChar);
        } // DoVisitSpecial

        protected override void DoVisitBreak(IRtfVisualBreak visualBreak)
        {
            if (!EnterVisual(visualBreak))
                return;

            switch (visualBreak.BreakKind)
            {
                case RtfVisualBreakKind.Line:
                    RenderLineBreak();
                    break;
                case RtfVisualBreakKind.Page:
                    break;
                case RtfVisualBreakKind.Paragraph:
                    if (IsInParagraph)
                    {
                        EndParagraph(); // close paragraph
                    }
                    else if (IsInListItem)
                    {
                        EndParagraph();
                        RenderEndTag(true); // close list item
                    }
                    else
                    {
                        BeginParagraph();
                        Writer.Write(NonBreakingSpace);
                        EndParagraph();
                    }
                    break;
                case RtfVisualBreakKind.Section:
                    break;
            }

            LeaveVisual(visualBreak);
        } // DoVisitBreak
        #endregion // RtfVisuals
    }
}