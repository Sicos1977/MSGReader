using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Documents;
using System.Xml;
using DataFormats = System.Windows.DataFormats;
using RichTextBox = System.Windows.Controls.RichTextBox;

/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader.Outlook
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
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                var convertThread = new Thread(Convert);
                convertThread.SetApartmentState(ApartmentState.STA);
                convertThread.Start();
                convertThread.Join();
            }
            else
                Convert();

            return _convertedRtf;
        }
        #endregion

        #region CleanUpCharacters
        /// <summary>
        /// Remove junk from the <paramref name="buffer"/>
        /// </summary>
        /// <param name="length">The length of the buffer</param>
        /// <param name="buffer">The bufer</param>
        private static void CleanUpCharacters(int length, IList<char> buffer)
        {
            for (var i = 0; i < length; i++)
            {
                var ch = buffer[i];
                int chi = ch;

                switch (chi)
                {
                    case 0: // embedded null
                    case 0x2000: // en quad
                    case 0x2001: // em quad
                    case 0x2002: // en space
                    case 0x2003: // em space
                    case 0x2004: // three-per-em space
                    case 0x2005: // four-per-em space
                    case 0x2006: // six-per-em space
                    case 0x2007: // figure space
                    case 0x2008: // puctuation space
                    case 0x2009: // thin space
                    case 0x200A: // hair space
                    case 0x200B: // zero-width space
                    case 0x200C: // zero-width non-joiner
                    case 0x200D: // zero-width joiner
                    case 0x202f: // no-break space
                    case 0x3000: // ideographic space
                    case 0x000C: // page break char
                    case 0x00A0: // non breaking space
                    case 0x00B6: // pilcro
                    case 0x2028: // line seperator
                    case 0x2029: // paragraph seperator
                        buffer[i] = ' ';
                        break;

                    case 0x00AD: // soft-hyphen
                    case 0x00B7: // middle dot
                    case 0x2010: // hyphen
                    case 0x2011: // non-breaking hyphen
                    case 0x2012: // figure dash
                    case 0x2013: // en dash
                    case 0x2014: // em dash
                    case 0x2015: // quote dash
                    case 0x2027: // hyphenation point
                    case 0x2043: // hyphen bullet
                    case 0x208B: // subscript minus
                    case 0xFE31: // vertical em dash
                    case 0xFE32: // vertical en dash
                    case 0xFE58: // small em dash
                    case 0xFE63: // small hyphen minus
                        buffer[i] = '-';
                        break;

                    case 0x00B0: // degree
                    case 0x2018: // left single quote
                    case 0x2019: // right single quote
                    case 0x201A: // low right single quote
                    case 0x201B: // high left single quote
                    case 0x2032: // prime
                    case 0x2035: // reversed prime
                    case 0x2039: // left-pointing angle quotation mark
                    case 0x203A: // right-pointing angle quotation mark
                        buffer[i] = '\'';
                        break;

                    case 0x201C: // left double quote
                    case 0x201D: // right double quote
                    case 0x201E: // low right double quote
                    case 0x201F: // high left double quote
                    case 0x2033: // double prime
                    case 0x2034: // triple prime
                    case 0x2036: // reversed double prime
                    case 0x2037: // reversed triple prime
                    case 0x00AB: // left-pointing double angle quotation mark
                    case 0x00BB: // right-pointing double angle quotation mark
                    case 0x3003: // ditto mark
                    case 0x301D: // reversed double prime quotation mark
                    case 0x301E: // double prime quotation mark
                    case 0x301F: // low double prime quotation mark
                        buffer[i] = '\"';
                        break;

                    case 0x00A7: // section-sign
                    case 0x2020: // dagger
                    case 0x2021: // double-dagger
                    case 0x2022: // bullet
                    case 0x2023: // triangle bullet
                    case 0x203B: // reference mark
                    case 0xFE55: // small colon
                        buffer[i] = ':';
                        break;

                    case 0x2024: // one dot leader
                    case 0x2025: // two dot leader
                    case 0x2026: // elipsis
                    case 0x3002: // ideographic full stop
                    case 0xFE30: // two dot vertical leader
                    case 0xFE52: // small full stop
                        buffer[i] = '.';
                        break;

                    case 0x3001: // ideographic comma
                    case 0xFE50: // small comma
                    case 0xFE51: // small ideographic comma
                        buffer[i] = ',';
                        break;

                    case 0xFE54: // small semicolon
                        buffer[i] = ';';
                        break;

                    case 0x00A6: // broken-bar
                    case 0x2016: // double vertical line
                        buffer[i] = '|';
                        break;

                    case 0x2017: // double low line
                    case 0x203E: // overline
                    case 0x203F: // undertie
                    case 0x2040: // character tie
                    case 0xFE33: // vertical low line
                    case 0xFE49: // dashed overline
                    case 0xFE4A: // centerline overline
                    case 0xFE4D: // dashed low line
                    case 0xFE4E: // centerline low line
                        buffer[i] = '_';
                        break;

                    case 0x301C: // wave dash
                    case 0x3030: // wavy dash
                    case 0xFE34: // vertical wavy low line
                    case 0xFE4B: // wavy overline
                    case 0xFE4C: // double wavy overline
                    case 0xFE4F: // wavy low line
                        buffer[i] = '~';
                        break;

                    case 0x2038: // caret
                    case 0x2041: // caret insertion point
                        buffer[i] = ' ';
                        break;

                    case 0x2030: // per-mille
                    case 0x2031: // per-ten thousand
                    case 0xFE6A: // small per-cent
                        buffer[i] = '%';
                        break;

                    case 0xFE6B: // small commercial at
                        buffer[i] = '@';
                        break;

                    case 0x00A9: // copyright
                        buffer[i] = 'c';
                        break;

                    case 0x00B5: // micro
                        buffer[i] = 'u';
                        break;

                    case 0x00AE: // registered
                        buffer[i] = 'r';
                        break;

                    case 0x207A: // superscript plus
                    case 0x208A: // subscript plus
                    case 0xFE62: // small plus
                        buffer[i] = '+';
                        break;

                    case 0x2044: // fraction slash
                        buffer[i] = '/';
                        break;

                    case 0x2042: // asterism
                    case 0xFE61: // small asterisk
                        buffer[i] = '*';
                        break;
                    case 0x208C: // subscript equal
                    case 0xFE66: // small equal
                        buffer[i] = '=';
                        break;
                    case 0xFE68: // small reverse solidus
                        buffer[i] = '\\';
                        break;

                    case 0xFE5F: // small number sign
                        buffer[i] = '#';
                        break;

                    case 0xFE60: // small ampersand
                        buffer[i] = '&';
                        break;

                    case 0xFE69: // small dollar sign
                        buffer[i] = '$';
                        break;

                    case 0x2045: // left square bracket with quill
                    case 0x3010: // left black lenticular bracket
                    case 0x3016: // left white lenticular bracket
                    case 0x301A: // left white square bracket
                    case 0xFE3B: // vertical left lenticular bracket
                    case 0xFF41: // vertical left corner bracket
                    case 0xFF43: // vertical white left corner bracket
                        buffer[i] = '[';
                        break;

                    case 0x2046: // right square bracket with quill
                    case 0x3011: // right black lenticular bracket
                    case 0x3017: // right white lenticular bracket
                    case 0x301B: // right white square bracket
                    case 0xFE3C: // vertical right lenticular bracket
                    case 0xFF42: // vertical right corner bracket
                    case 0xFF44: // vertical white right corner bracket
                        buffer[i] = ']';
                        break;

                    case 0x208D: // subscript left parenthesis
                    case 0x3014: // left tortise-shell bracket
                    case 0x3018: // left white tortise-shell bracket
                    case 0xFE35: // vertical left parenthesis
                    case 0xFE39: // vertical left tortise-shell bracket
                    case 0xFE59: // small left parenthesis
                    case 0xFE5D: // small left tortise-shell bracket
                        buffer[i] = '(';
                        break;

                    case 0x208E: // subscript right parenthesis
                    case 0x3015: // right tortise-shell bracket
                    case 0x3019: // right white tortise-shell bracket
                    case 0xFE36: // vertical right parenthesis
                    case 0xFE3A: // vertical right tortise-shell bracket
                    case 0xFE5A: // small right parenthesis
                    case 0xFE5E: // small right tortise-shell bracket
                        buffer[i] = ')';
                        break;

                    case 0x3008: // left angle bracket
                    case 0x300A: // left double angle bracket
                    case 0xFF3D: // vertical left double angle bracket
                    case 0xFF3F: // vertical left angle bracket
                    case 0xFF64: // small less-than
                        buffer[i] = '<';
                        break;

                    case 0x3009: // right angle bracket
                    case 0x300B: // right double angle bracket
                    case 0xFF3E: // vertical right double angle bracket
                    case 0xFF40: // vertical right angle bracket
                    case 0xFF65: // small greater-than
                        buffer[i] = '>';
                        break;

                    case 0xFE37: // vertical left curly bracket
                    case 0xFE5B: // small left curly bracket
                        buffer[i] = '{';
                        break;

                    case 0xFE38: // vertical right curly bracket
                    case 0xFE5C: // small right curly bracket
                        buffer[i] = '}';
                        break;

                    case 0x00A1: // inverted exclamation mark
                    case 0x00AC: // not
                    case 0x203C: // double exclamation mark
                    case 0x203D: // interrobang
                    case 0xFE57: // small exclamation mark
                        buffer[i] = '!';
                        break;

                    case 0x00BF: // inverted question mark
                    case 0xFE56: // small question mark
                        buffer[i] = '?';
                        break;

                    case 0x00B9: // superscript one
                        buffer[i] = '1';
                        break;

                    case 0x00B2: // superscript two
                        buffer[i] = '2';
                        break;

                    case 0x00B3: // superscript three
                        buffer[i] = '3';
                        break;

                    case 0x2070: // superscript zero
                    case 0x2074: // superscript four
                    case 0x2075: // superscript five
                    case 0x2076: // superscript six
                    case 0x2077: // superscript seven
                    case 0x2078: // superscript eight
                    case 0x2079: // superscript nine
                    case 0x2080: // subscript zero
                    case 0x2081: // subscript one
                    case 0x2082: // subscript two
                    case 0x2083: // subscript three
                    case 0x2084: // subscript four
                    case 0x2085: // subscript five
                    case 0x2086: // subscript six
                    case 0x2087: // subscript seven
                    case 0x2088: // subscript eight
                    case 0x2089: // subscript nine
                    case 0x3021: // Hangzhou numeral one
                    case 0x3022: // Hangzhou numeral two
                    case 0x3023: // Hangzhou numeral three
                    case 0x3024: // Hangzhou numeral four
                    case 0x3025: // Hangzhou numeral five
                    case 0x3026: // Hangzhou numeral six
                    case 0x3027: // Hangzhou numeral seven
                    case 0x3028: // Hangzhou numeral eight
                    case 0x3029: // Hangzhou numeral nine
                        chi = chi & 0x000F;
                        buffer[i] = System.Convert.ToChar(chi);
                        break;

                    // ONE is at ZERO location... careful
                    case 0x3220: // parenthesized ideograph one
                    case 0x3221: // parenthesized ideograph two
                    case 0x3222: // parenthesized ideograph three
                    case 0x3223: // parenthesized ideograph four
                    case 0x3224: // parenthesized ideograph five
                    case 0x3225: // parenthesized ideograph six
                    case 0x3226: // parenthesized ideograph seven
                    case 0x3227: // parenthesized ideograph eight
                    case 0x3228: // parenthesized ideograph nine
                    case 0x3280: // circled ideograph one
                    case 0x3281: // circled ideograph two
                    case 0x3282: // circled ideograph three
                    case 0x3283: // circled ideograph four
                    case 0x3284: // circled ideograph five
                    case 0x3285: // circled ideograph six
                    case 0x3286: // circled ideograph seven
                    case 0x3287: // circled ideograph eight
                    case 0x3288: // circled ideograph nine
                        chi = (chi & 0x000F) + 1;
                        buffer[i] = System.Convert.ToChar(chi);
                        break;

                    case 0x3007: // ideographic number zero
                    case 0x24EA: // circled number zero
                        buffer[i] = '0';
                        break;

                    default:
                        if (0xFF01 <= ch // fullwidth exclamation mark 
                            && ch <= 0xFF5E) // fullwidth tilde
                        {
                            // the fullwidths line up with ASCII low subset
                            buffer[i] = System.Convert.ToChar(chi & 0xFF00 + '!' - 1);
                            //ch = ch & 0xFF00 + '!' - 1;               
                        }
                        else if (0x2460 <= ch // circled one
                                 && ch <= 0x2468) // circled nine
                        {
                            buffer[i] = System.Convert.ToChar(chi - 0x2460 + '1');
                            //ch = ch - 0x2460 + '1';
                        }
                        else if (0x2474 <= ch // parenthesized one
                                 && ch <= 0x247C) // parenthesized nine
                        {
                            buffer[i] = System.Convert.ToChar(chi - 0x2474 + '1');
                            // ch = ch - 0x2474 + '1';
                        }
                        else if (0x2488 <= ch // one full stop
                                 && ch <= 0x2490) // nine full stop
                        {
                            buffer[i] = System.Convert.ToChar(chi - 0x2488 + '1');
                            //ch = ch - 0x2488 + '1';
                        }
                        else if (0x249C <= ch // parenthesized small a
                                 && ch <= 0x24B5) // parenthesized small z
                        {
                            buffer[i] = System.Convert.ToChar(chi - 0x249C + 'a');
                            //ch = ch - 0x249C + 'a';
                        }
                        else if (0x24B6 <= ch // circled capital A
                                 && ch <= 0x24CF) // circled capital Z
                        {
                            buffer[i] = System.Convert.ToChar(chi - 0x24B6 + 'A');
                            //ch = ch - 0x24B6 + 'A';
                        }
                        else if (0x24D0 <= ch // circled small a
                                 && ch <= 0x24E9) // circled small z
                        {
                            buffer[i] = System.Convert.ToChar(chi - 0x24D0 + 'a');
                            //ch = ch - 0x24D0 + 'a';
                        }
                        else if (0x2500 <= ch // box drawing (begin)
                                 && ch <= 0x257F) // box drawing (end)
                        {
                            buffer[i] = '|';
                        }
                        else if (0x2580 <= ch // block elements (begin)
                                 && ch <= 0x259F) // block elements (end)
                        {
                            buffer[i] = '#';
                        }
                        else if (0x25A0 <= ch // geometric shapes (begin)
                                 && ch <= 0x25FF) // geometric shapes (end)
                        {
                            buffer[i] = '*';
                        }
                        else if (0x2600 <= ch // dingbats (begin)
                                 && ch <= 0x267F) // dingbats (end)
                        {
                            buffer[i] = '.';
                        }
                        break;
                }
            }
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
            try
            {
                using (var rtfMemoryStream = new MemoryStream())
                using (var rtfStreamWriter = new StreamWriter(rtfMemoryStream))
                {
                    rtfStreamWriter.Write(_rtf);
                    rtfStreamWriter.Flush();
                    rtfMemoryStream.Seek(0, SeekOrigin.Begin);

                    // Load the MemoryStream into TextRange ranging from start to end of RichTextBox.
                    textRange.Load(rtfMemoryStream, DataFormats.Rtf);
                }
            }
            catch (ArgumentException)
            {
                // Sometimes we have malformed RTF, in this case we try to clean it and convert it again
                using (var rtfMemoryStream = new MemoryStream())
                using (var rtfStreamWriter = new StreamWriter(rtfMemoryStream))
                {
                    var rtfCharArray = _rtf.ToCharArray();
                    CleanUpCharacters(rtfCharArray.Length, rtfCharArray);
                    _rtf = new string(rtfCharArray);
                    rtfStreamWriter.Write(_rtf);
                    rtfStreamWriter.Flush();
                    rtfMemoryStream.Seek(0, SeekOrigin.Begin);
                    // Load the MemoryStream into TextRange ranging from start to end of RichTextBox.
                    textRange.Load(rtfMemoryStream, DataFormats.Rtf);
                }
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
                // Root FlowDocument element is missing
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
        private static void WriteFormattingProperties(XmlReader xamlReader, XmlWriter htmlWriter, StringBuilder inlineStyle)
        {
            if (xamlReader == null) throw new ArgumentNullException("xamlReader");

            // Clear string builder for the inline style
            inlineStyle.Remove(0, inlineStyle.Length);

            if (!xamlReader.HasAttributes)
                return;

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
                    case "CellSpacing":
                        css = "border-spacing: " + xamlReader.Value + ";";
                        break;

                    case "Width":
                        css = "width:" + xamlReader.Value + "px; height:auto;";
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
        private static string ParseXamlColor(string color)
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
        private static string ParseXamlThickness(string thickness)
        {
            var values = thickness.Split(',');

            for (var i = 0; i < values.Length; i++)
            {
                double value;
                if (double.TryParse(values[i], NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value))
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

                                htmlWriter.WriteString(xamlReader.Value.Replace("\t", "\u00A0\u00A0\u00A0\u00A0"));
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
        ///     Converts an element notation of complex property into
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
                    case "LineBreak":
                        htmlElementName = "br";
                        break;

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
        private static bool ReadNextToken(XmlReader xamlReader)
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
