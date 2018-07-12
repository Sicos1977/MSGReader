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