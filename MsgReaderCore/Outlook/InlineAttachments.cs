//
// InlineAttachment.cs
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
    /// <summary>
    /// Used as a temporary placeholder for information about the inline attachents
    /// </summary>
    internal sealed class InlineAttachment
    {
        #region Properties
        /// <summary>
        /// Returns the rendering position for the attachmnt
        /// </summary>
        public int? RenderingPosition { get; }

        /// <summary>
        /// Returns the name of the icon when this attachment is part of an RTF body and is
        /// shown as an icon
        /// </summary>
        public string IconFileName { get; }

        /// <summary>
        /// Returns the name for the attachment
        /// </summary>
        public string AttachmentFileName { get; }

        /// <summary>
        /// Returns the full name for the attachment
        /// </summary>
        public string FullName { get; }
        #endregion

        #region Constructors
        public InlineAttachment(int renderingPosition,
            string attachmentFileName)
        {
            RenderingPosition = renderingPosition;
            AttachmentFileName = attachmentFileName;
            FullName = attachmentFileName;
        }

        public InlineAttachment(string iconFileName,
            string attachmentFileName,
            string fullName)
        {
            IconFileName = iconFileName;
            AttachmentFileName = attachmentFileName;
            FullName = fullName;
        }
        #endregion
    }
}
