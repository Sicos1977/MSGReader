using System.Collections;
using System.IO;
using System.Text;

namespace MsgReader.Outlook
{
    #region RecipientType
    /// <summary>
    /// The <see cref="RecipientRow.DisplayType"/>
    /// </summary>
    public enum RecipientType
    {
        /// <summary>
        /// No type is set
        /// </summary>
        NoType = 0x0,
        /// <summary>
        /// X500DN
        /// </summary>
        
        X500Dn = 0x1,
        /// <summary>
        /// Ms mail
        /// </summary>
        MsMail = 0x2,

        /// <summary>
        /// SMTP
        /// </summary>
        Smtp = 0x3,

        /// <summary>
        /// Fax
        /// </summary>
        Fax = 0x4,
        
        /// <summary>
        /// Professional office system
        /// </summary>
        ProfessionalOfficeSystem = 0x5,

        /// <summary>
        /// Personal distribution list 1
        /// </summary>
        PersonalDistributionList1 = 0x6,

        /// <summary>
        /// Personal distribution list 2
        /// </summary>
        PersonalDistributionList2 = 0x7
    }
    #endregion
    
    #region DisplayType
    /// <summary>
    /// An enumeration. This field MUST be present when the Type field
    /// of the RecipientFlags field is set to X500DN(0x1) and MUST NOT be present otherwise.This
    /// value specifies the display type of this address.Valid values for this field are specified in the
    /// following table.
    /// </summary>
    public enum DisplayType
    {
        /// <summary>
        /// A messaging user
        /// </summary>
        MessagingUser = 0x00,

        /// <summary>
        /// A distribution list
        /// </summary>
        DistributionList = 0x01,

        /// <summary>
        /// A forum, such as a bulletin board service or a public or shared folder
        /// </summary>
        Forum = 0x02,

        /// <summary>
        /// An automated agent
        /// </summary>
        AutomatedAgent = 0x03,

        /// <summary>
        /// An Address Book object defined for a large group, such as helpdesk, accounting, coordinator, or
        /// department
        /// </summary>
        AddressBook = 0x04,

        /// <summary>
        /// A private, personally administered distribution list
        /// </summary>
        PrivateDistributionList = 0x05,

        /// <summary>
        /// An Address Book object known to be from a foreign or remote messaging system
        /// </summary>
        RemoteAddressBook = 0x06
    }
    #endregion

    /// <summary>
    ///     The PidLidAppointmentUnsendableRecipients  property ([MS-OXPROPS] section 2.35) contains a list of 
    ///     unsendable attendees. This property is not required but SHOULD be set
    /// </summary>
    public class UnsendableRecipients
    {
        #region Properties
        /// <summary>
        /// An integer that specifies the number of structures in the RecipientRow field
        /// </summary>
        public uint RowCount { get; internal set; }
        #endregion

        internal UnsendableRecipients(byte[] data)
        {
            var binaryReader = new BinaryReader(new MemoryStream(data));
            RowCount = binaryReader.ReadUInt32();
        }
    }

    /// <summary>
    ///     An array of RecipientRow structures, as specified in [MS-OXCDATA] section 2.8.3. 
    ///     Each structure specifies an unsendable attendee. The RowCount field specifies the 
    ///     number of structures contained in this field. For details about properties that can 
    ///     be set on RecipientRow structures for Calendar objects and meeting-related objects, 
    ///     see section 2.2.4.10. 
    /// </summary>
    public class RecipientRow
    {
        #region Properties
        /// <summary>
        /// The <see cref="RecipientType"/>
        /// </summary>
        public RecipientType RecipientType { get; private set; }

        /// <summary>
        /// The address prefix used
        /// </summary>
        public string AddressPrefixUsed { get; private set; }

        /// <summary>
        /// The <see cref="DisplayType"/>
        /// </summary>
        public DisplayType DisplayType { get; private set; }
        
        /// <summary>
        /// Returns null when not available
        /// </summary>
        public string X500Dn { get; private set; }
        #endregion

        #region Constructor
        internal RecipientRow(BinaryReader binaryReader)
        {
            var b = new BitArray(binaryReader.ReadBytes(4));
            
            // If this flag is b'1', a different transport is responsible for delivery to this recipient(1).
            var r = b[0];
            
            // If this flag is b'1', the value of the TransmittableDisplayName field is the same as the value of the DisplayName field.
            var s = b[1];
            
            // If this flag is b'1', the TransmittableDisplayName (section 2.8.3.2) field is included.
            var t = b[2];
            
            // If this flag is b'1', the DisplayName (section 2.8.3.2) field is included.
            var d = b[3];
            
            // If this flag is b'1', the EmailAddress (section 2.8.3.2) field is included.
            var e = b[4];
            
            // This enumeration specifies the type of address.
            var bt = new BitArray(3);
            bt.Set(0, b[5]);
            bt.Set(1, b[6]);
            bt.Set(2, b[7]);
            var array = new int[1];
            bt.CopyTo(array, 0);
            RecipientType = (RecipientType)array[0];
            
            // If this flag is b'1', this recipient (1) has a non-standard address type and the AddressType field is included.
            var o = b[8];

            // Reserved (4 bits): (mask 0x7800) The server MUST set this to b'0000'.

            //  If this flag is b'1', the SimpleDisplayName field is included.
            var i = b[13];
            
            // If this flag is b'1', the associated string properties are in Unicode with a 2-
            // byte terminating null character; if this flag is b'0', string properties are MBCS with a single
            // terminating null character, in the code page sent to the server in the EcDoConnectEx method,
            // as specified in [MS-OXCRPC] section 3.1.4.1, or the Connect request type<6>, as specified in
            // [MS-OXCMAPIHTTP] section 2.2.4.1.
            var u = b[14];
            
            // If b'1', this flag specifies that the recipient (1) does not support receiving
            // rich text messages.
            var n = b[15];

            if (RecipientType == RecipientType.X500Dn)
                X500Dn = ReadNullTerminatedString(binaryReader);

            if (RecipientType == RecipientType.PersonalDistributionList1 ||
                RecipientType == RecipientType.PersonalDistributionList2)
                X500Dn = ReadNullTerminatedString(binaryReader);
        }
        #endregion

        #region ReadNullTerminatedString
        /// <summary>
        /// Reads from the <see cref="binaryReader"/> until a null terminated char is read
        /// </summary>
        /// <param name="binaryReader">The <see cref="BinaryReader"/></param>
        /// <returns></returns>
        private string ReadNullTerminatedString(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            var chr = binaryReader.ReadChar();
            while (chr != 0)
            {
                stringBuilder.Append(chr);
                chr = binaryReader.ReadChar();
            }

            return stringBuilder.ToString();
        }
        #endregion
    }
}
