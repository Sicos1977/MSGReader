//
// AddresBookEntryId.cs
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
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System.IO;
using MsgReader.Helpers;

namespace MsgReader.Outlook
{
    /// <summary>
    ///     An Address Book EntryID structure specifies several types of Address Book objects, including
    ///     individual users, distribution lists, containers, and templates.
    /// </summary>
    /// <remarks>
    ///     See https://msdn.microsoft.com/en-us/library/ee160588(v=exchg.80).aspx
    /// </remarks>
    public class AddressBookEntryId
    {
        #region Properties
        // what circumstances a short-term EntryID is valid. However, in any EntryID stored in a 
        // property value, these 4 bytes MUST be zero, indicating a long-term EntryID.
        /// <summary>
        ///     Flags (4 bytes): This value MUST be set to 0x00000000. Bits in this field indicate under
        /// </summary>
        public byte[] Flags { get; private set; }

        /// <summary>
        ///     The X500 DN of the Address Book object.
        /// </summary>
        /// <remarks>
        ///     A distinguished name (DN), in Teletex form, of an object that is in an address book. An X500 DN can be more limited
        ///     in the size and number of relative distinguished names (RDNs) than a full DN.
        /// </remarks>
        public string X500Dn { get; private set; }
        #endregion

        #region Constructor
        internal AddressBookEntryId(BinaryReader binaryReader)
        {
            Flags = binaryReader.ReadBytes(4);
            X500Dn = Strings.ReadNullTerminatedString(binaryReader, false);
        }
        #endregion
    }
}