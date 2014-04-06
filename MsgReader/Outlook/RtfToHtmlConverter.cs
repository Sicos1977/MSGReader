using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Xml;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    /// <summary>
    /// This class is used to convert RTF to HTML by using the RichTextEditBox
    /// </summary>
    internal class RtfToHtmlConverter
    {
        #region Fields
        /// <summary>
        /// The RTF string that needs to be converted
        /// </summary>
        private string _rtf;

        /// <summary>
        /// The RTF string that is converted to HTML
        /// </summary>
        private string _convertedRtf;
        #endregion

        #region ConvertRtfToHtml
        /// <summary>
        /// Convert RTF to HTML by using the Windows RichTextBox
        /// </summary>
        /// <param name="rtf">The rtf string</param>
        /// <returns></returns>
        public string ConvertRtfToHtml(string rtf)
        {
            // Because the RichtTextBox is a control that needs to run in STA mode we always start
            // a thread in STA mode
            _rtf = rtf;
            var convertThread = new Thread(Convert);
            convertThread.SetApartmentState(ApartmentState.STA);
            convertThread.Start();
            convertThread.Join();
            return _convertedRtf;
        }
        #endregion

        #region Convert
        /// <summary>
        /// Do the actual conversion by using a RichTextBox
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2202:Do not dispose objects multiple times")]
        private void Convert()
        {
            var richTextBox = new RichTextBox();
            if (string.IsNullOrEmpty(_rtf))
            {
                _convertedRtf = string.Empty;
                return;
            }

            var textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);

            // Create a MemoryStream of the Rtf content
            using (var rtfMemoryStream = new MemoryStream())
            using (var rtfStreamWriter = new StreamWriter(rtfMemoryStream))
            {
                rtfStreamWriter.Write(_rtf);
                rtfStreamWriter.Flush();
                rtfMemoryStream.Seek(0, SeekOrigin.Begin);

                // Load the MemoryStream into TextRange ranging from start to end of RichTextBox.
                textRange.Load(rtfMemoryStream, DataFormats.Rtf);
            }

            using (var rtfMemoryStream = new MemoryStream())
            {
                textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                textRange.Save(rtfMemoryStream, DataFormats.Xaml);
                rtfMemoryStream.Seek(0, SeekOrigin.Begin);
                using (var rtfStreamReader = new StreamReader(rtfMemoryStream))
                    _convertedRtf = ConvertXamlToHtml(rtfStreamReader.ReadToEnd());
            }
        }
        #endregion

        #region ConvertXamlToHtml
        /// <summary>
        ///     Converts an xaml into html
        /// </summary>
        /// <param name="xamlString"></param>
        /// <returns></returns>
        private string ConvertXamlToHtml(string xamlString)
        {
            var xamlReader = new XmlTextReader(new StringReader("<FlowDocument>" + xamlString + "</FlowDocument>"));

            var htmlStringBuilder = new StringBuilder(100);
            var htmlWriter = new XmlTextWriter(new StringWriter(htmlStringBuilder));

            if (!WriteFlowDocument(xamlReader, htmlWriter))
                return string.Empty;

            var htmlString = htmlStringBuilder.ToString();

            return htmlString;
        }
        #endregion

        #region WriteFlowDocument
        /// <summary>
        /// Processes a root level element of XAML (normally it's FlowDocument element).
        /// </summary>
        /// <param name="xamlReader">XmlTextReader for a source xaml</param>
        /// <param name="htmlWriter">XmlTextWriter producing resulting html</param>
        private bool WriteFlowDocument(XmlTextReader xamlReader, XmlTextWriter htmlWriter)
        {
            if (!ReadNextToken(xamlReader))
                // Xaml content is empty - nothing to convert
                return false;

            if (xamlReader.NodeType != XmlNodeType.Element || xamlReader.Name != "FlowDocument")
                // Root FlowDocument elemet is missing
                return false;

            // Create a buffer StringBuilder for collecting css properties for inline STYLE attributes
            // on every element level (it will be re-initialized on every level).
            var inlineStyle = new StringBuilder();

            htmlWriter.WriteStartElement("html");
            htmlWriter.WriteStartElement("body");

            WriteFormattingProperties(xamlReader, htmlWriter, inlineStyle);
            WriteElementContent(xamlReader, htmlWriter, inlineStyle);

            htmlWriter.WriteEndElement();
            htmlWriter.WriteEndElement();
            return true;
        }
        #endregion

        #region WriteFormattingProperties
        /// <summary>
        ///     Reads attributes of the current xaml element and converts them into appropriate html attributes or css styles
        /// </summary>
        /// <param name="xamlReader">
        ///     XmlTextReader which is expected to be at XmlNodeType.Element (opening element tag) position.
        ///     The reader will remain at the same level after function complete.</param>
        /// <param name="htmlWriter">
        ///     XmlTextWriter for output html, which is expected to be in after WriteStartElement state.
        /// </param>
        /// <param name="inlineStyle">
        ///     String builder for collecting css properties for inline STYLE attribute.
        /// </param>
        private void WriteFormattingProperties(XmlReader xamlReader, XmlWriter htmlWriter, StringBuilder inlineStyle)
        {
            if (xamlReader == null) throw new ArgumentNullException("xamlReader");

            // Clear string builder for the inline style
            inlineStyle.Remove(0, inlineStyle.Length);

            if (!xamlReader.HasAttributes)
            {
                return;
            }

            var borderSet = false;

            while (xamlReader.MoveToNextAttribute())
            {
                string css = null;

                switch (xamlReader.Name)
                {
                    // Character fomatting properties
                    // ------------------------------
                    case "Background":
                        css = "background-color:" + ParseXamlColor(xamlReader.Value) + ";";
                        break;

                    case "FontFamily":
                        css = "font-family:" + xamlReader.Value + ";";
                        break;

                    case "FontStyle":
                        css = "font-style:" + xamlReader.Value.ToLower() + ";";
                        break;

                    case "FontWeight":
                        css = "font-weight:" + xamlReader.Value.ToLower() + ";";
                        break;

                    case "FontStretch":
                        break;

                    case "FontSize":
                        css = "font-size:" + xamlReader.Value + ";";
                        break;

                    case "Foreground":
                        css = "color:" + ParseXamlColor(xamlReader.Value) + ";";
                        break;

                    case "TextDecorations":
                        css = xamlReader.Value.ToLower() == "strikethrough"
                            ? "text-decoration:line-through;"
                            : "text-decoration:underline;";
                        break;

                    case "TextEffects":
                        break;

                    case "Emphasis":
                        break;

                    case "StandardLigatures":
                        break;

                    case "Variants":
                        break;

                    case "Capitals":
                        break;

                    case "Fraction":
                        break;

                    // Paragraph formatting properties
                    // -------------------------------
                    case "Padding":
                        css = "padding:" + ParseXamlThickness(xamlReader.Value) + ";";
                        break;

                    case "Margin":
                        css = "margin:" + ParseXamlThickness(xamlReader.Value) + ";";
                        break;

                    case "BorderThickness":
                        css = "border-width:" + ParseXamlThickness(xamlReader.Value) + ";";
                        borderSet = true;
                        break;

                    case "BorderBrush":
                        css = "border-color:" + ParseXamlColor(xamlReader.Value) + ";";
                        borderSet = true;
                        break;

                    case "LineHeight":
                        break;

                    case "TextIndent":
                        css = "text-indent:" + xamlReader.Value + ";";
                        break;

                    case "TextAlignment":
                        css = "text-align:" + xamlReader.Value + ";";
                        break;

                    case "IsKeptTogether":
                        break;

                    case "IsKeptWithNext":
                        break;

                    case "ColumnBreakBefore":
                        break;

                    case "PageBreakBefore":
                        break;

                    case "FlowDirection":
                        break;

                    // Table attributes
                    // ----------------
                    case "Width":
                        css = "width:" + xamlReader.Value + ";";
                        break;

                    case "ColumnSpan":
                        htmlWriter.WriteAttributeString("colspan", xamlReader.Value);
                        break;

                    case "RowSpan":
                        htmlWriter.WriteAttributeString("rowspan", xamlReader.Value);
                        break;

                    // Hyperlink Attributes
                    case "NavigateUri":
                        htmlWriter.WriteAttributeString("href", xamlReader.Value);
                        break;

                    case "TargetName":
                        htmlWriter.WriteAttributeString("target", xamlReader.Value);
                        break;
                }

                if (css != null)
                    inlineStyle.Append(css);
            }

            if (borderSet)
                inlineStyle.Append("border-style:solid;mso-element:para-border-div;");

            // Return the xamlReader back to element level
            xamlReader.MoveToElement();
        }
        #endregion

        #region ParseXamlColor
        /// <summary>
        /// Parses xaml color
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        private string ParseXamlColor(string color)
        {
            if (color.StartsWith("#"))
                // Remove transparancy value
                color = "#" + color.Substring(3);

            return color;
        }
        #endregion

        #region ParseXamlThickness
        /// <summary>
        /// Parses xaml thickness
        /// </summary>
        /// <param name="thickness"></param>
        /// <returns></returns>
        private string ParseXamlThickness(string thickness)
        {
            var values = thickness.Split(',');

            for (var i = 0; i < values.Length; i++)
            {
                double value;
                if (double.TryParse(values[i], out value))
                    values[i] = Math.Ceiling(value).ToString(CultureInfo.InvariantCulture);
                else
                    values[i] = "1";
            }

            string cssThickness;

            switch (values.Length)
            {
                case 1:
                    cssThickness = thickness;
                    break;

                case 2:
                    cssThickness = values[1] + " " + values[0];
                    break;

                case 4:
                    cssThickness = values[1] + " " + values[2] + " " + values[3] + " " + values[0];
                    break;

                default:
                    cssThickness = values[0];
                    break;
            }

            return cssThickness;
        }
        #endregion

        #region WriteElementContent
        /// <summary>
        ///     Reads a content of current xaml element, converts it
        /// </summary>
        /// <param name="xamlReader">
        ///     XmlTextReader which is expected to be at XmlNodeType.Element (opening element tag) position.
        /// </param>
        /// <param name="htmlWriter">
        ///     May be null, in which case we are skipping the xaml element; witout producing any output to html.
        /// </param>
        /// <param name="inlineStyle">
        ///     StringBuilder used for collecting css properties for inline STYLE attribute.
        /// </param>
        private void WriteElementContent(XmlTextReader xamlReader, XmlTextWriter htmlWriter, StringBuilder inlineStyle)
        {
            var elementContentStarted = false;

            if (xamlReader.IsEmptyElement)
            {
                if (htmlWriter == null || inlineStyle.Length <= 0) return;

                // Output STYLE attribute and clear inlineStyle buffer.
                htmlWriter.WriteAttributeString("style", inlineStyle.ToString());
                inlineStyle.Remove(0, inlineStyle.Length);
            }
            else
            {
                while (ReadNextToken(xamlReader) && xamlReader.NodeType != XmlNodeType.EndElement)
                {
                    switch (xamlReader.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (xamlReader.Name.Contains("."))
                                AddComplexProperty(xamlReader, inlineStyle);
                            else
                            {
                                if (htmlWriter != null && !elementContentStarted && inlineStyle.Length > 0)
                                {
                                    // Output STYLE attribute and clear inlineStyle buffer.
                                    htmlWriter.WriteAttributeString("style", inlineStyle.ToString());
                                    inlineStyle.Remove(0, inlineStyle.Length);
                                }
                                elementContentStarted = true;
                                WriteElement(xamlReader, htmlWriter, inlineStyle);
                            }
                            break;

                        case XmlNodeType.Comment:
                            if (htmlWriter != null)
                            {
                                if (!elementContentStarted && inlineStyle.Length > 0)
                                    htmlWriter.WriteAttributeString("style", inlineStyle.ToString());

                                htmlWriter.WriteComment(xamlReader.Value);
                            }
                            elementContentStarted = true;
                            break;

                        case XmlNodeType.CDATA:
                        case XmlNodeType.Text:
                        case XmlNodeType.SignificantWhitespace:
                            if (htmlWriter != null)
                            {
                                if (!elementContentStarted && inlineStyle.Length > 0)
                                    htmlWriter.WriteAttributeString("style", inlineStyle.ToString());

                                htmlWriter.WriteString(xamlReader.Value);
                            }

                            elementContentStarted = true;
                            break;
                    }
                }
            }
        }
        #endregion

        #region AddComplexProperty
        /// <summary>
        ///     Conberts an element notation of complex property into
        /// </summary>
        /// <param name="xamlReader">
        ///     On entry this XmlTextReader must be on Element start tag; on exit - on EndElement tag.
        /// </param>
        /// <param name="inlineStyle">
        ///     StringBuilder containing a value for STYLE attribute.
        /// </param>
        private void AddComplexProperty(XmlTextReader xamlReader, StringBuilder inlineStyle)
        {
            if (inlineStyle != null && xamlReader.Name.EndsWith(".TextDecorations"))
                inlineStyle.Append("text-decoration:underline;");

            // Skip the element representing the complex property
            WriteElementContent(xamlReader, null, null);
        }
        #endregion

        #region WriteElement
        /// <summary>
        ///     Converts a xaml element into an appropriate html element.
        /// </summary>
        /// <param name="xamlReader">
        ///     On entry this XmlTextReader must be on Element start tag;
        ///     on exit - on EndElement tag.
        /// </param>
        /// <param name="htmlWriter">
        ///     May be null, in which case we are skipping xaml content
        ///     without producing any html output
        /// </param>
        /// <param name="inlineStyle">
        ///     StringBuilder used for collecting css properties for inline STYLE attributes on every level.
        /// </param>
        private void WriteElement(XmlTextReader xamlReader, XmlTextWriter htmlWriter, StringBuilder inlineStyle)
        {
            if (htmlWriter == null)
                // Skipping mode; recurse into the xaml element without any output
                WriteElementContent(xamlReader, null, null);
            else
            {
                string htmlElementName = null;

                switch (xamlReader.Name)
                {
                    case "Run":
                    case "Span":
                        htmlElementName = "span";
                        break;

                    case "InlineUIContainer":
                        htmlElementName = "span";
                        break;

                    case "Bold":
                        htmlElementName = "b";
                        break;

                    case "Italic":
                        htmlElementName = "i";
                        break;

                    case "Paragraph":
                        htmlElementName = "p";
                        break;

                    case "BlockUIContainer":
                        htmlElementName = "div";
                        break;

                    case "Section":
                        htmlElementName = "div";
                        break;

                    case "Table":
                        htmlElementName = "table";
                        break;

                    case "TableColumn":
                        htmlElementName = "col";
                        break;

                    case "TableRowGroup":
                        htmlElementName = "tbody";
                        break;

                    case "TableRow":
                        htmlElementName = "tr";
                        break;

                    case "TableCell":
                        htmlElementName = "td";
                        break;

                    case "List":
                        var marker = xamlReader.GetAttribute("MarkerStyle");
                        if (marker == null || marker == "None" || marker == "Disc" || marker == "Circle" || marker == "Square" || marker == "Box")
                            htmlElementName = "ul";
                        else
                            htmlElementName = "OL";
                        break;

                    case "ListItem":
                        htmlElementName = "li";
                        break;

                    case "Hyperlink":
                        htmlElementName = "a";
                        break;
                }

                if (htmlElementName != null)
                {
                    htmlWriter.WriteStartElement(htmlElementName);
                    WriteFormattingProperties(xamlReader, htmlWriter, inlineStyle);
                    WriteElementContent(xamlReader, htmlWriter, inlineStyle);
                    htmlWriter.WriteEndElement();
                }
                else
                {
                    // Skip this unrecognized xaml element
                    WriteElementContent(xamlReader, null, null);
                }
            }
        }
        #endregion

        #region ReadNextToken
        /// <summary>
        ///     Reads several items from xamlReader skipping all non-significant stuff.
        /// </summary>
        /// <param name="xamlReader">
        ///     XmlTextReader from tokens are being read.
        /// </param>
        /// <returns>
        ///     True if new token is available; false if end of stream reached.
        /// </returns>
        private bool ReadNextToken(XmlReader xamlReader)
        {
            while (xamlReader.Read())
            {
                switch (xamlReader.NodeType)
                {
                    case XmlNodeType.Element:
                    case XmlNodeType.EndElement:
                    case XmlNodeType.None:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.Text:
                    case XmlNodeType.SignificantWhitespace:
                        return true;

                    case XmlNodeType.Whitespace:
                        if (xamlReader.XmlSpace == XmlSpace.Preserve)
                            return true;

                        // Ignore insignificant whitespace
                        break;

                    case XmlNodeType.EndEntity:
                    case XmlNodeType.EntityReference:
                        break;

                    case XmlNodeType.Comment:
                        return true;
                }
            }
            return false;
        }
        #endregion
    }
}
