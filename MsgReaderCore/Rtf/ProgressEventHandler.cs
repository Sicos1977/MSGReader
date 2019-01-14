//
// ProgressEventHandler.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
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

namespace MsgReader.Rtf
{
    #region Delegates
    /// <summary>
    /// Progress event handler type
    /// </summary>
    /// <param name="sender">sender</param>
    /// <param name="e">event arguments</param>
    internal delegate void ProgressEventHandler(object sender , ProgressEventArgs e);
    #endregion

    /// <summary>
    /// Progress event arguments
    /// </summary>
    internal class ProgressEventArgs : EventArgs
    {
        #region Properties
        /// <summary>
        /// Progress max value
        /// </summary>
        public int MaxValue { get; }

        /// <summary>
        /// Current value
        /// </summary>
        public int Value { get; }

        /// <summary>
        /// Progress message
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Cancel operation
        /// </summary>
        public bool Cancel { get; set; }
        #endregion

        #region Constructor
        public ProgressEventArgs(int max, int value, string message)
        {
            Cancel = false;
            MaxValue = max;
            Value = value;
            Message = message;
        }
        #endregion
    }
}
