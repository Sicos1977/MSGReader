using System;
using MsgReader.Outlook;

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

namespace MsgReader.Exceptions
{
    /// <summary>
    /// Raised when it is not possible to read the <see cref="Storage.Attachment"/> from
    /// the <see cref="Storage.Message"/>
    /// </summary>
    public class MRUnknownAttachmentFormat : Exception
    {
        internal MRUnknownAttachmentFormat()
        {
        }

        internal MRUnknownAttachmentFormat(string message)
            : base(message)
        {
        }

        internal MRUnknownAttachmentFormat(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}