using System;
using System.Collections.Generic;
using System.Net;
using System.Text;

namespace MsgViewer.Helpers
{
    /// <summary>
    ///     Converts HTML to plain text.
    /// </summary>
    public class HtmlToText
    {
        #region Internal class TextBuilder
        /// <summary>
        ///     A StringBuilder class that helps eliminate excess whitespace.
        /// </summary>
        internal class TextBuilder
        {
            private readonly StringBuilder _currLine;
            private readonly StringBuilder _text;
            private int _emptyLines;
            private bool _preformatted;

            // Construction
            public TextBuilder()
            {
                _text = new StringBuilder();
                _currLine = new StringBuilder();
                _emptyLines = 0;
                _preformatted = false;
            }

            /// <summary>
            ///     Normally, extra whitespace characters are discarded.
            ///     If this property is set to true, they are passed
            ///     through unchanged.
            /// </summary>
            public bool Preformatted
            {
                get { return _preformatted; }
                set
                {
                    if (value)
                    {
                        // Clear line buffer if changing to
                        // preformatted mode
                        if (_currLine.Length > 0)
                            FlushCurrLine();
                        _emptyLines = 0;
                    }
                    _preformatted = value;
                }
            }

            /// <summary>
            ///     Clears all current text.
            /// </summary>
            public void Clear()
            {
                _text.Length = 0;
                _currLine.Length = 0;
                _emptyLines = 0;
            }

            /// <summary>
            ///     Writes the given string to the output buffer.
            /// </summary>
            /// <param name="s"></param>
            public void Write(string s)
            {
                foreach (var c in s)
                    Write(c);
            }

            /// <summary>
            ///     Writes the given character to the output buffer.
            /// </summary>
            /// <param name="c">Character to write</param>
            public void Write(char c)
            {
                if (_preformatted)
                {
                    // Write preformatted character
                    _text.Append(c);
                }
                else
                {
                    switch (c)
                    {
                        case '\r':
                            break;
                        case '\n':
                            FlushCurrLine();
                            break;
                        default:
                            if (Char.IsWhiteSpace(c))
                            {
                                // Write single space character
                                var len = _currLine.Length;
                                if (len == 0 || !Char.IsWhiteSpace(_currLine[len - 1]))
                                    _currLine.Append(' ');
                            }
                            else
                            {
                                // Add character to current line
                                _currLine.Append(c);
                            }
                            break;
                    }
                }
            }

            // Appends the current line to output buffer
            protected void FlushCurrLine()
            {
                // Get current line
                var line = _currLine.ToString().Trim();

                // Determine if line contains non-space characters
                var tmp = line.Replace("&nbsp;", String.Empty);
                if (tmp.Length == 0)
                {
                    // An empty line
                    _emptyLines++;
                    if (_emptyLines < 2 && _text.Length > 0)
                        _text.AppendLine(line);
                }
                else
                {
                    // A non-empty line
                    _emptyLines = 0;
                    _text.AppendLine(line);
                }

                // Reset current line
                _currLine.Length = 0;
            }

            /// <summary>
            ///     Returns the current output as a string.
            /// </summary>
            public override string ToString()
            {
                if (_currLine.Length > 0)
                    FlushCurrLine();
                return _text.ToString();
            }
        }
        #endregion

        #region Fields
        private static Dictionary<string, string> _tags;
        private static HashSet<string> _ignoreTags;
        private string _html;
        private int _pos;
        private TextBuilder _text;
        #endregion

        #region Constructor
        public HtmlToText()
        {
            _tags = new Dictionary<string, string>
            {
                {"address", "\n"},
                {"blockquote", "\n"},
                {"div", "\n"},
                {"dl", "\n"},
                {"fieldset", "\n"},
                {"form", "\n"},
                {"h1", "\n"},
                {"/h1", "\n"},
                {"h2", "\n"},
                {"/h2", "\n"},
                {"h3", "\n"},
                {"/h3", "\n"},
                {"h4", "\n"},
                {"/h4", "\n"},
                {"h5", "\n"},
                {"/h5", "\n"},
                {"h6", "\n"},
                {"/h6", "\n"},
                {"p", "\n"},
                {"/p", "\n"},
                {"table", "\n"},
                {"/table", "\n"},
                {"ul", "\n"},
                {"/ul", "\n"},
                {"ol", "\n"},
                {"/ol", "\n"},
                {"/li", "\n"},
                {"br", "\n"},
                {"/td", "\t"},
                {"/tr", "\n"},
                {"/pre", "\n"}
            };

            _ignoreTags = new HashSet<string> {"script", "noscript", "style", "object"};
        }
        #endregion

        #region EndOfText
        /// <summary>
        /// Returns true when we are at the end of te text
        /// </summary>
        protected bool EndOfText
        {
            get { return (_pos >= _html.Length); }
        }
        #endregion

        #region Convert
        /// <summary>
        ///     Converts the given HTML to plain text and returns the result.
        /// </summary>
        /// <param name="html">HTML to be converted</param>
        /// <returns>Resulting plain text</returns>
        public string Convert(string html)
        {
            // Initialize state variables
            _text = new TextBuilder();
            _html = html;
            _pos = 0;

            // Process input
            while (!EndOfText)
            {
                if (Peek() == '<')
                {
                    // HTML tag
                    bool selfClosing;
                    var tag = ParseTag(out selfClosing);

                    // Handle special tag cases
                    if (tag == "body")
                    {
                        // Discard content before <body>
                        _text.Clear();
                    }
                    else if (tag == "/body")
                    {
                        // Discard content after </body>
                        _pos = _html.Length;
                    }
                    else if (tag == "pre")
                    {
                        // Enter preformatted mode
                        _text.Preformatted = true;
                        EatWhitespaceToNextLine();
                    }
                    else if (tag == "/pre")
                    {
                        // Exit preformatted mode
                        _text.Preformatted = false;
                    }

                    string value;
                    if (_tags.TryGetValue(tag, out value))
                        _text.Write(value);

                    if (_ignoreTags.Contains(tag))
                        EatInnerContent(tag);
                }
                else if (Char.IsWhiteSpace(Peek()))
                {
                    // Whitespace (treat all as space)
                    _text.Write(_text.Preformatted ? Peek() : ' ');
                    MoveAhead();
                }
                else
                {
                    // Other text
                    _text.Write(Peek());
                    MoveAhead();
                }
            }
            // Return result
            return WebUtility.HtmlDecode(_text.ToString());
        }
        #endregion

        #region ParseTag
        /// <summary>
        ///     Eats all characters that are part of the current tag and returns information about that tag
        /// </summary>
        /// <param name="selfClosing"></param>
        /// <returns></returns>
        private string ParseTag(out bool selfClosing)
        {
            var tag = String.Empty;
            selfClosing = false;

            if (Peek() != '<') return tag;
            MoveAhead();

            // Parse tag name
            EatWhitespace();
            var start = _pos;
            if (Peek() == '/')
                MoveAhead();
            while (!EndOfText && !Char.IsWhiteSpace(Peek()) &&
                   Peek() != '/' && Peek() != '>')
                MoveAhead();
            tag = _html.Substring(start, _pos - start).ToLower();

            // Parse rest of tag
            while (!EndOfText && Peek() != '>')
            {
                if (Peek() == '"' || Peek() == '\'')
                    EatQuotedValue();
                else
                {
                    if (Peek() == '/')
                        selfClosing = true;
                    MoveAhead();
                }
            }
            MoveAhead();
            return tag;
        }
        #endregion

        #region EatInnerContent
        /// <summary>
        /// Consumes inner content from the current tag
        /// </summary>
        /// <param name="tag"></param>
        private void EatInnerContent(string tag)
        {
            var endTag = "/" + tag;

            while (!EndOfText)
            {
                if (Peek() == '<')
                {
                    // Consume a tag
                    bool selfClosing;
                    if (ParseTag(out selfClosing) == endTag)
                        return;
                    // Use recursion to consume nested tags
                    if (!selfClosing && !tag.StartsWith("/"))
                        EatInnerContent(tag);
                }
                else
                    MoveAhead();
            }
        }
        #endregion

        #region Helpers
        /// <summary>
        ///     Safely returns the character at the current position
        /// </summary>
        /// <returns></returns>
        private char Peek()
        {
            return (_pos < _html.Length) ? _html[_pos] : (char) 0;
        }

        /// <summary>
        ///     Safely advances to current position to the next character
        /// </summary>
        private void MoveAhead()
        {
            _pos = Math.Min(_pos + 1, _html.Length);
        }

        /// <summary>
        ///     Moves the current position to the next non-whitespace character
        /// </summary>
        private void EatWhitespace()
        {
            while (Char.IsWhiteSpace(Peek()))
                MoveAhead();
        }

        /// <summary>
        /// Moves the current position to the next non-whitespace
        /// character or the start of the next line, whichever
        /// comes first
        /// </summary>
        private void EatWhitespaceToNextLine()
        {
            while (Char.IsWhiteSpace(Peek()))
            {
                var c = Peek();
                MoveAhead();
                if (c == '\n')
                    break;
            }
        }

        /// <summary>
        /// Moves the current position past a quoted value
        /// </summary>
        private void EatQuotedValue()
        {
            var c = Peek();
            if (c != '"' && c != '\'')
                return;

            // Opening quote
            MoveAhead();
            // Find end of value
            _pos = _html.IndexOfAny(new[] {c, '\r', '\n'}, _pos);
            if (_pos < 0)
                _pos = _html.Length;
            else
                MoveAhead(); // Closing quote
        }
        #endregion
    }
}