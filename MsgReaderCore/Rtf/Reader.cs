//
// Reader.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2022 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.Collections.Generic;
using System.IO;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf reader
    /// </summary>
    internal sealed class Reader : IDisposable
    {
        #region Fields
        private readonly Stack<LayerInfo> _layerStack = new Stack<LayerInfo>();
        private readonly Lex _lex;
        #endregion

        #region Properties
        public TextReader InnerReader { get; private set; }

        /// <summary>
        /// Current token
        /// </summary>
        public Token CurrentToken { get; private set; }

        /// <summary>
        /// Current token's type
        /// </summary>
        public RtfTokenType TokenType => CurrentToken?.Type ?? RtfTokenType.None;

        /// <summary>
        /// Current keyword
        /// </summary>
        public string Keyword => CurrentToken?.Key;

        /// <summary>
        /// If current token has a parameter
        /// </summary>
        public bool HasParam => CurrentToken != null && CurrentToken.HasParam;

        /// <summary>
        /// Current parameter
        /// </summary>
        public int Parameter => CurrentToken?.Param ?? 0;

        /// <summary>
        /// Lost token
        /// </summary>
        public Token LastToken { get; private set; }

        /// <summary>
        /// When set to <c>true</c> then we are parsing an RTF unicode
        /// high - low surrogate
        /// </summary>
        internal bool ParsingHighLowSurrogate { get; set; }

        /// <summary>
        /// When <see cref="ParsingHighLowSurrogate"/> is set to <c>true</c>
        /// then this will containt the high surrogate value when we are
        /// parsing the low surrogate value
        /// </summary>
        internal int? HighSurrogateValue { get; set; }

        /// <summary>
        /// <see cref="LayerInfo"/>
        /// </summary>
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
        /// Initialize instance from text reader
        /// </summary>
        public Reader(TextReader reader)
        {
            CurrentToken = null;
            InnerReader = reader;
            _lex = new Lex(InnerReader);
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

        #region Close
        /// <summary>
        /// Close the inner reader
        /// </summary>
        private void Close()
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
        private void DefaultProcess()
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
            LastToken = CurrentToken;

            CurrentToken = _lex.NextToken();
            if (CurrentToken == null || CurrentToken.Type == RtfTokenType.Eof)
            {
                CurrentToken = null;
                return null;
            }

            switch (CurrentToken.Type)
            {
                case RtfTokenType.GroupStart when _layerStack.Count == 0:
                    _layerStack.Push(new LayerInfo());
                    break;
              
                case RtfTokenType.GroupStart:
                {
                    var info = _layerStack.Peek();
                    _layerStack.Push(info.Clone());
                    break;
                }
                
                case RtfTokenType.GroupEnd:
                {
                    if (_layerStack.Count > 0)
                        _layerStack.Pop();
                    break;
                }
            }

            DefaultProcess();

            return CurrentToken;
        }
        #endregion

        #region ReadToEndOfGroup
        /// <summary>
        /// Read and ignore data , until just the end of the current group, preserve the end.
        /// </summary>
        public void ReadToEndOfGroup()
        {
            var level = 0;

            while (true)
            {
                var c = InnerReader.Peek();
                if (c == -1)
                    break;
                
                if (c == '{')
                    level++;
                
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
    }
}