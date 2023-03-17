//
// Task.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2023 Magic-Sessions. (www.magic-sessions.com)
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

namespace MsgReader.Outlook;

public partial class Storage
{
    /// <summary>
    ///     Class used to contain all the log document information
    /// </summary>
    public sealed class Log : Storage
    {
        #region Properties
        /// <summary>
        ///     Returns <c>true</c> when the <see cref="Storage.Log" /> document has been printed
        /// </summary>
        public bool? DocumentPrinted { get; }

        /// <summary>
        ///     Returns <c>true</c> when the <see cref="Storage.Log" /> document has been saved
        /// </summary>
        public bool? DocumentSaved { get; }

        /// <summary>
        ///     Returns <c>true</c> when the <see cref="Storage.Log" /> document has been routed
        /// </summary>
        public bool? DocumentRouted { get; }

        /// <summary>
        ///     Returns <c>true</c> when the <see cref="Storage.Log" /> document has been posted
        /// </summary>
        public bool? DocumentPosted { get; }

        /// <summary>
        ///     Returns the log type
        /// </summary>
        public string Type { get; }

        /// <summary>
        ///     Returns the log type description
        /// </summary>
        public string TypeDescription { get; }

        /// <summary>
        ///     Returns the start date/time of the log
        /// </summary>
        public DateTime? Start { get; }

        /// <summary>
        ///     Returns the end date/time of the log
        /// </summary>
        public DateTime? End { get; }

        /// <summary>
        ///     Returns the duration of the log
        /// </summary>
        public long? Duration { get; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Initializes a new instance of the <see cref="Storage.Task" /> class.
        /// </summary>
        /// <param name="message"> The message. </param>
        internal Log(Storage message) : base(message._rootStorage)
        {
            _namedProperties = message._namedProperties;
            _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

            Type = GetMapiPropertyString(MapiTags.LogType);
            TypeDescription = GetMapiPropertyString(MapiTags.LogTypeDesc);
            DocumentPrinted = GetMapiPropertyBool(MapiTags.LogDocPrinted);
            DocumentSaved = GetMapiPropertyBool(MapiTags.LogDocSaved);
            DocumentRouted = GetMapiPropertyBool(MapiTags.LogDocRouted);
            DocumentPosted = GetMapiPropertyBool(MapiTags.LogDocPosted);
            Start = GetMapiPropertyDateTime(MapiTags.LogStart);
            End = GetMapiPropertyDateTime(MapiTags.LogEnd);
            Duration = GetMapiPropertyInt32(MapiTags.LogDuration);
        }
        #endregion
    }
}