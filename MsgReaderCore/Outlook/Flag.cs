//
// Flag.cs
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

namespace MsgReader.Outlook
{
    #region Public enum FlagStatus
    /// <summary>
    /// The flag state of an message
    /// </summary>
    public enum FlagStatus
    {
        /// <summary>
        /// The msg object has been flagged as completed
        /// </summary>
        Complete = 1,

        /// <summary>
        /// The msg object has been flagged and marked as a task
        /// </summary>
        Marked = 2
    }
    #endregion

    public partial class Storage
    {
        /// <summary>
        /// Class used to contain all the flag (follow up) information of a <see cref="Storage.Message"/>.
        /// </summary>
        public sealed class Flag : Storage
        {
            #region Properties
            /// <summary>
            /// Returns the flag request text
            /// </summary>
            public string Request { get; }

            /// <summary>
            /// Returns the <see cref="FlagStatus">Status</see> of the flag
            /// </summary>
            public FlagStatus? Status { get; }
            #endregion

            #region Constructor
            /// <summary>
            ///   Initializes a new instance of the <see cref="Storage.Flag" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Flag(Storage message) : base(message._rootStorage)
            {
                _namedProperties = message._namedProperties;
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

                Request = GetMapiPropertyString(MapiTags.FlagRequest);

                var status = GetMapiPropertyInt32(MapiTags.PR_FLAG_STATUS);
                if (status == null)
                    Status = null;
                else
                    Status = (FlagStatus)status; 
            }
            #endregion
        }
    }
}