using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf plain text container
    /// </summary>
    internal class TextContainer
    {
        #region Fields
        private readonly ByteBuffer _byteBuffer = new ByteBuffer();
        private StringBuilder _stringBuilder = new StringBuilder();
        #endregion

        #region Properties
        /// <summary>
        /// Owner document
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public DomDocument Document { get; set; }

        /// <summary>
        /// Check if this container has some text
        /// </summary>
        public bool HasContent
        {
            get
            {
                CheckBuffer();
                return _stringBuilder.Length > 0;
            }
        }

        /// <summary>
        /// text value
        /// </summary>
        public string Text
        {
            get
            {
                CheckBuffer();
                return _stringBuilder.ToString();
            }
        }

        public int Level { get; set; }
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize instance
        /// </summary>
        /// <param name="document">owner document</param>
        public TextContainer(DomDocument document)
        {
            Level = 0;
            Document = document;
        }
        #endregion

        #region Append
        /// <summary>
        /// Append text content
        /// </summary>
        /// <param name="text"></param>
        public void Append(string text)
        {
            if (string.IsNullOrEmpty(text) == false)
            {
                CheckBuffer();
                _stringBuilder.Append(text);
            }
        }
        #endregion

        #region Accept
        /// <summary>
        /// Accept rtf token
        /// </summary>
        /// <param name="token">RTF token</param>
        /// <param name="reader"></param>
        /// <returns>Is accept it?</returns>
        public bool Accept(Token token, Reader reader)
        {
            if (token == null)
                return false;
            
            if (token.Type == RtfTokenType.Text)
            {
                if (reader != null)
                {
                    if (token.Key[0] == '?')
                    {
                        if (reader.LastToken != null)
                        {
                            if (reader.LastToken.Type == RtfTokenType.Keyword
                                && reader.LastToken.Key == "u"
                                && reader.LastToken.HasParam)
                            {
                                if (token.Key.Length > 0)
                                    CheckBuffer();
                                return true;
                            }
                        }
                    }
                }
                CheckBuffer();
                _stringBuilder.Append(token.Key);
                return true;
            }

            if (token.Type == RtfTokenType.Control && token.Key == "'" && token.HasParam)
            {
                if (reader.CurrentLayerInfo.CheckUcValueCount())
                    _byteBuffer.Add((byte) token.Param);
                return true;
            }

            if (token.Key == Consts.U && token.HasParam)
            {
                // Unicode char
                CheckBuffer();
                _stringBuilder.Append((char) token.Param);
                reader.CurrentLayerInfo.UCValueCount = reader.CurrentLayerInfo.UcValue;
                return true;
            }

            if (token.Key == Consts.Tab)
            {
                CheckBuffer();
                _stringBuilder.Append("\t");
                return true;
            }

            if (token.Key == Consts.Emdash)
            {
                CheckBuffer();
                _stringBuilder.Append('-');
                return true;
            }

            if (token.Key == "")
            {
                CheckBuffer();
                _stringBuilder.Append('-');
                return true;
            }
            return false;
        }
        #endregion

        #region Clear
        /// <summary>
        /// Clear value
        /// </summary>
        public void Clear()
        {
            _byteBuffer.Clear();
            _stringBuilder = new StringBuilder();
        }
        #endregion

        #region CheckBuffer
        /// <summary>
        /// Check if buffer still containts any text
        /// </summary>
        private void CheckBuffer()
        {
            if (_byteBuffer.Count > 0)
            {
                var text = _byteBuffer.GetString(Document.RuntimeEncoding);
                _stringBuilder.Append(text);
                _byteBuffer.Clear();
            }
        }
        #endregion
    }
}