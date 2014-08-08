using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtw writer, use it to write a new Rtf document
    /// </summary>
    internal sealed class Writer : IDisposable
    {
        #region Consts
        /// <summary>
        /// Hex characters
        /// </summary>
        private const string Hexs = "0123456789abcdef";
        #endregion

        #region Fields
        /// <summary>
        /// Inner text writer
        /// </summary>
        private TextWriter _textWriter;

        /// <summary>
        /// Text encoding
        /// </summary>
        // ReSharper disable once UnusedMember.Local
        private Encoding _unicode = Encoding.Unicode;

        /// <summary>
        ///     current line head
        /// </summary>
        // ReSharper disable once NotAccessedField.Local
        private int _lineHead;

        /// <summary>
        /// Current position
        /// </summary>
        private int _position;
        #endregion

        #region Properties
        public Encoding Encoding { get; set; }

        /// <summary>
        /// Whether output rtf code with indent style
        /// </summary>
        /// <remarks>
        /// In rtf formation , recommend do not use indent . this option just to debugger ,
        /// in software development , use this option can genereate indented rtf code friendly to read,
        /// but after debug , recommend clear this option and set this attribute = false.
        /// </remarks>
        public bool Indent { get; set; }

        /// <summary>
        /// String used to indent
        /// </summary>
        public string IndentString { get; set; }

        /// <summary>
        /// Level of nested groups
        /// </summary>
        public int GroupLevel { get; private set; }
        #endregion

        #region RTFWriter
        /// <summary>
        /// Initialize instance
        /// </summary>
        /// <param name="textWriter">text writer</param>
        public Writer(TextWriter textWriter)
        {
            IndentString = "   ";
            Encoding = Encoding.Default;
            _textWriter = textWriter;
        }

        /// <summary>
        /// Initialize instance
        /// </summary>
        /// <param name="fileName">File name</param>
        public Writer(string fileName)
        {
            IndentString = "   ";
            Encoding = Encoding.Default;
            _textWriter = new StreamWriter(fileName, false, Encoding.ASCII);
        }
        #endregion

        #region Dispose
        public void Dispose()
        {
            Close();
        }
        #endregion

        #region Close
        public void Close()
        {
            if (GroupLevel > 0)
                throw new Exception("One or more groups did not finish");

            if (_textWriter != null)
            {
                _textWriter.Close();
                _textWriter = null;
            }
        }
        #endregion

        #region Flush
        public void Flush()
        {
            if (_textWriter != null)
                _textWriter.Flush();
        }
        #endregion

        #region WriteGroup
        /// <summary>
        /// Write completed group wich one keyword
        /// </summary>
        /// <param name="keyWord">keyword</param>
        public void WriteGroup(string keyWord)
        {
            WriteStartGroup();
            WriteKeyword(keyWord);
            WriteEndGroup();
        }

        /// <summary>
        /// Begin write group
        /// </summary>
        public void WriteStartGroup()
        {
            if (Indent)
            {
                InnerWriteNewLine();
                _textWriter.Write("{");
            }
            else
            {
                _textWriter.Write("{");
            }
            GroupLevel ++;
        }

        /// <summary>
        /// End write group
        /// </summary>
        public void WriteEndGroup()
        {
            GroupLevel--;
            if (GroupLevel < 0)
                throw new Exception("Group level error");

            if (Indent)
            {
                InnerWriteNewLine();
                InnerWrite("}");
            }
            else
                InnerWrite("}");
        }
        #endregion

        #region WriteRaw
        /// <summary>
        /// Write raw text
        /// </summary>
        /// <param name="text">text</param>
        public void WriteRaw(string text)
        {
            if (!string.IsNullOrEmpty(text))
                InnerWrite(text);
        }
        #endregion

        #region WriteKeyWord
        /// <summary>
        /// Write keyword
        /// </summary>
        /// <param name="keyword">keyword</param>
        public void WriteKeyword(string keyword)
        {
            WriteKeyword(keyword, false);
        }

        /// <summary>
        /// Write keyword
        /// </summary>
        /// <param name="keyword">keyword</param>
        /// <param name="external">True if it is an external keyword</param>
        public void WriteKeyword(string keyword, bool external)
        {
            if (string.IsNullOrEmpty(keyword))
                throw new ArgumentNullException("keyword");

            if (Indent == false && (keyword == "par" || keyword == "pard"))
                InnerWrite(Environment.NewLine);

            if (Indent)
            {
                if (keyword == "par" || keyword == "pard")
                    InnerWriteNewLine();
            }

            InnerWrite(external ? "\\*\\" : "\\");
            InnerWrite(keyword);
        }
        #endregion

        #region WriteText
        /// <summary>
        /// Write plain text
        /// </summary>
        /// <param name="text">Text</param>
        public void WriteText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return;

            WriteText(text, true);
        }

        public void WriteUnicodeText(string text)
        {
            if (string.IsNullOrEmpty(text) == false)
            {
                WriteKeyword("uc1");
                foreach (var c in text)
                {
                    if (c > 127)
                    {
                        int v = c;
                        var v2 = (short) v;
                        WriteKeyword("u" + v2);
                        WriteRaw(" ?");
                    }
                    else
                    {
                        InnerWriteChar(c);
                    }
                }
            }
        }

        /// <summary>
        /// Write plain text, with the option to automaticly add a white space character
        /// </summary>
        /// <param name="text">Text</param>
        /// <param name="autoAddWhitespace">Write a white space automatic</param>
        public void WriteText(string text, bool autoAddWhitespace)
        {
            if (string.IsNullOrEmpty(text))
                return;

            if (autoAddWhitespace)
                InnerWrite(' ');

            foreach (var c in text)
                InnerWriteChar(c);
        }

        private void InnerWriteChar(char c)
        {
            if (c == '\t')
            {
                WriteKeyword("tab");
                InnerWrite(' ');
            }
            if (c > 32 && c < 127)
            {
                // Some special characters , must be converted
                if (c == '\\' || c == '{' || c == '}')
                    InnerWrite('\\');
                InnerWrite(c);
            }
            else
            {
                var bytes = Encoding.GetBytes(c.ToString(CultureInfo.InvariantCulture));
                foreach (var b in bytes)
                {
                    InnerWrite("\\\'");
                    WriteByte(b);
                }
            }
        }
        #endregion

        #region WriteBytes
        /// <summary>
        /// Write binary data
        /// </summary>
        /// <param name="bytes">Binary data</param>
        public void WriteBytes(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
                return;

            WriteRaw(" ");
            
            for (var count = 0; count < bytes.Length; count ++)
            {
                if ((count%32) == 0)
                {
                    WriteRaw(Environment.NewLine);
                    WriteIndent();
                }
                else if ((count%8) == 0)
                {
                    WriteRaw(" ");
                }
                var b = bytes[count];
                var h = (b & 0xf0) >> 4;
                var l = b & 0xf;
                _textWriter.Write(Hexs[h]);
                _textWriter.Write(Hexs[l]);
                _position += 2;
            }
        }

        /// <summary>
        /// Write a byte
        /// </summary>
        /// <param name="b">byte</param>
        public void WriteByte(byte b)
        {
            var h = (b & 0xf0) >> 4;
            var l = b & 0xf;
            _textWriter.Write(Hexs[h]);
            _textWriter.Write(Hexs[l]);
            _position += 2;
        }
        #endregion

        #region Internal functions
        private void InnerWrite(char c)
        {
            _position ++;
            _textWriter.Write(c);
        }

        private void InnerWrite(string text)
        {
            _position += text.Length;
            _textWriter.Write(text);
        }

        private void InnerWriteNewLine()
        {
            if (Indent)
            {
                if (_position > 0)
                {
                    InnerWrite(Environment.NewLine);
                    _lineHead = _position;
                    WriteIndent();
                }
            }
        }

        private void WriteIndent()
        {
            if (Indent)
            {
                for (var count = 0; count < GroupLevel; count ++)
                {
                    InnerWrite(IndentString);
                }
            }
        }
        #endregion
    }
}