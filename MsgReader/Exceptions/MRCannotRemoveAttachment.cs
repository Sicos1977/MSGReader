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
    /// Raised when it is not possible to remove the <see cref="Storage.Attachment"/> or <see cref="Storage.Message"/> from
    /// the <see cref="Storage.Message"/>
    /// </summary>
    public class MRCannotRemoveAttachment : Exception
    {
        internal MRCannotRemoveAttachment()
        {
        }

        internal MRCannotRemoveAttachment(string message)
            : base(message)
        {
        }

        internal MRCannotRemoveAttachment(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}