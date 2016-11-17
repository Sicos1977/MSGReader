using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf reader
    /// </summary>
    internal sealed class Reader : IDisposable
    {
        #region Fields
        private readonly Stack<LayerInfo> _layerStack = new Stack<LayerInfo>();
        private bool _firstTokenInGroup;
        private Stream _stream;
        private Lex _lex;
        #endregion

        #region Properties
        public TextReader InnerReader { get; private set; }

        /// <summary>
        /// Current token
        /// </summary>
        public Token CurrentToken { get; private set; }

        /// <summary>
        /// The <see cref="DocumentFormatInfo"/>
        /// </summary>
        public DocumentFormatInfo Format { get; internal set; }

        /// <summary>
        /// Current token's type
        /// </summary>
        public RtfTokenType TokenType
        {
            get { return CurrentToken == null ? RtfTokenType.None : CurrentToken.Type; }
        }

        /// <summary>
        /// Current keyword
        /// </summary>
        public string Keyword
        {
            get { return CurrentToken == null ? null : CurrentToken.Key; }
        }

        /// <summary>
        /// If current token has a parameter
        /// </summary>
        public bool HasParam
        {
            get { return CurrentToken != null && CurrentToken.HasParam; }
        }

        /// <summary>
        /// Current parameter
        /// </summary>
        public int Parameter
        {
            get { return CurrentToken == null ? 0 : CurrentToken.Param; }
        }

        public int ContentPosition
        {
            get
            {
                if (_stream == null)
                    return 0;
                return (int) _stream.Position;
            }
        }

        public int ContentLength
        {
            get
            {
                if (_stream == null)
                    return 0;
                return (int)_stream.Length;
            }
        }

        /// <summary>
        /// Current token is the first token in owner group
        /// </summary>
        public bool FirstTokenInGroup
        {
            get { return _firstTokenInGroup; }
        }

        /// <summary>
        /// Lost token
        /// </summary>
        public Token LastToken { get; private set; }

        public int Level { get; private set; }

        /// <summary>
        /// Total of this object handle tokens
        /// </summary>
        public int TokenCount { get; set; }

        // ReSharper disable once MemberCanBePrivate.Global
        public bool EnableDefaultProcess { get; set; }

        public LayerInfo CurrentLayerInfo
        {
            get
            {
                if (_layerStack.Count == 0)
                    _layerStack.Push(new LayerInfo());
                return _layerStack.Peek();
            }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Initialize instance
        /// </summary>
        public Reader()
        {
            EnableDefaultProcess = true;
        }

        /// <summary>
        /// Initialize instance from file
        /// </summary>
        // ReSharper disable once UnusedMember.Global
        public Reader(string fileName)
        {
            EnableDefaultProcess = true;
            LoadRTFFile(fileName);
        }

        /// <summary>
        /// Initialize instance from stream
        /// </summary>
        public Reader(Stream stream)
        {
            EnableDefaultProcess = true;
            var reader = new StreamReader(stream, Encoding.ASCII);
            LoadReader(reader);
            _stream = stream;
        }

        /// <summary>
        /// Initialize instance from text reader
        /// </summary>
        public Reader(TextReader reader)
        {
            EnableDefaultProcess = true;
            LoadReader(reader);
        }
        #endregion

        #region Displose
        /// <summary>
        /// Dispose this object
        /// </summary>
        public void Dispose()
        {
            Close();
        }
        #endregion

        #region Load
        /// <summary>
        /// Load rtf from file
        /// </summary>
        /// <param name="fileName">spcial file name</param>
        /// <returns>is operation successful</returns>
        public bool LoadRTFFile(string fileName)
        {
            CurrentToken = null;
            if (File.Exists(fileName))
            {
                var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                InnerReader = new StreamReader(stream, Encoding.ASCII);
                _stream = stream;
                _lex = new Lex(InnerReader);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Load rtf from reader
        /// </summary>
        /// <param name="reader">text reader</param>
        /// <returns>is operation successful</returns>
        public void LoadReader(TextReader reader)
        {
            //.Clear();
            CurrentToken = null;
            InnerReader = reader;
            _lex = new Lex(InnerReader);
        }

        /// <summary>
        /// Load rtf from string
        /// </summary>
        /// <param name="text">RTF text</param>
        /// <returns>is operation successful</returns>
        public bool LoadRTFText(string text)
        {
            //myTokenStack.Clear();
            CurrentToken = null;
            if (text != null && text.Length > 3)
            {
                InnerReader = new StringReader(text);
                _lex = new Lex(InnerReader);
                return true;
            }
            return false;
        }
        #endregion

        #region Close
        /// <summary>
        /// Close the inner reader
        /// </summary>
        public void Close()
        {
            if (InnerReader != null)
            {
                InnerReader.Close();
                InnerReader = null;
            }
        }
        #endregion

        #region PeekTokenType
        /// <summary>
        /// Get next token type
        /// </summary>
        /// <returns></returns>
        public RtfTokenType PeekTokenType()
        {
            return _lex.PeekTokenType();
        }
        #endregion

        #region DefaultProcess
        public void DefaultProcess()
        {
            if (CurrentToken == null) return;
            switch (CurrentToken.Key)
            {
                case "uc":
                    CurrentLayerInfo.UcValue = Parameter;
                    break;
				case "u":
		            if (InnerReader.Peek() == '?')
			            InnerReader.Read();
		            break;
            }
        }
        #endregion

        #region ReadToken
        /// <summary>
        /// Read token
        /// </summary>
        /// <returns>token read</returns>
        public Token ReadToken()
        {
            _firstTokenInGroup = false;
            LastToken = CurrentToken;

            if (LastToken != null && LastToken.Type == RtfTokenType.GroupStart)
                _firstTokenInGroup = true;
            
            CurrentToken = _lex.NextToken();
            if (CurrentToken == null || CurrentToken.Type == RtfTokenType.Eof)
            {
                CurrentToken = null;
                return null;
            }

            TokenCount++;

            if (CurrentToken.Type == RtfTokenType.GroupStart)
            {
                if (_layerStack.Count == 0)
                    _layerStack.Push(new LayerInfo());
                else
                {
                    var info = _layerStack.Peek();
                    _layerStack.Push(info.Clone());
                }
                Level++;
            }
            else if (CurrentToken.Type == RtfTokenType.GroupEnd)
            {
                if (_layerStack.Count > 0)
                    _layerStack.Pop();
                Level--;
            }

            if (EnableDefaultProcess)
                DefaultProcess();

            return CurrentToken;
        }
        #endregion

        #region ReadToEndGround
        /// <summary>
        /// Read and ignore data , until just the end of current group, preserve the end.
        /// </summary>
        public void ReadToEndGround()
        {
            var level = 0;
            while (true)
            {
                var c = InnerReader.Peek();
                if (c == -1)
                {
                    break;
                }
                if (c == '{')
                {
                    level++;
                }
                else if (c == '}')
                {
                    level--;
                    if (level < 0)
                    {
                        break;
                    }
                }
                InnerReader.Read();
            }
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            return "RTFReader Level:" + Level + " " + Keyword;
        }
        #endregion
    }
}