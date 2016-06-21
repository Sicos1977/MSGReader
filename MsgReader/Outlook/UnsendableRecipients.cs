using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using MsgReader.Helpers;

namespace MsgReader.Outlook
{
    #region Enum RecipientType
    /// <summary>
    ///     The <see cref="RecipientRow.DisplayType" />
    /// </summary>
    public enum RecipientType
    {
        /// <summary>
        ///     No type is set
        /// </summary>
        NoType = 0x0,

        /// <summary>
        ///     X500DN
        /// </summary>
        X500Dn = 0x1,

        /// <summary>
        ///     Ms mail
        /// </summary>
        MsMail = 0x2,

        /// <summary>
        ///     SMTP
        /// </summary>
        Smtp = 0x3,

        /// <summary>
        ///     Fax
        /// </summary>
        Fax = 0x4,

        /// <summary>
        ///     Professional office system
        /// </summary>
        ProfessionalOfficeSystem = 0x5,

        /// <summary>
        ///     Personal distribution list 1
        /// </summary>
        PersonalDistributionList1 = 0x6,

        /// <summary>
        ///     Personal distribution list 2
        /// </summary>
        PersonalDistributionList2 = 0x7
    }
    #endregion

    #region Enum DisplayType
    /// <summary>
    ///     An enumeration. This field MUST be present when the Type field
    ///     of the RecipientFlags field is set to X500DN(0x1) and MUST NOT be present otherwise.This
    ///     value specifies the display type of this address.Valid values for this field are specified in the
    ///     following table.
    /// </summary>
    public enum DisplayType
    {
        /// <summary>
        ///     A messaging user
        /// </summary>
        MessagingUser = 0x00,

        /// <summary>
        ///     A distribution list
        /// </summary>
        DistributionList = 0x01,

        /// <summary>
        ///     A forum, such as a bulletin board service or a public or shared folder
        /// </summary>
        Forum = 0x02,

        /// <summary>
        ///     An automated agent
        /// </summary>
        AutomatedAgent = 0x03,

        /// <summary>
        ///     An Address Book object defined for a large group, such as helpdesk, accounting, coordinator, or
        ///     department
        /// </summary>
        AddressBook = 0x04,

        /// <summary>
        ///     A private, personally administered distribution list
        /// </summary>
        PrivateDistributionList = 0x05,

        /// <summary>
        ///     An Address Book object known to be from a foreign or remote messaging system
        /// </summary>
        RemoteAddressBook = 0x06
    }
    #endregion

    #region Enum AddressBookEntryId
    /// <summary>
    ///     An integer representing the type of the object. It MUST be one of the values from the following table.
    /// </summary>
    public enum AddressBookEntryIdType
    {
        LocalMailUser = 0x00000000,
        DistributionList = 0x00000001,
        BulletinBoardOrPublicFolder = 0x00000002,
        AutomatedMailBox = 0x00000003,
        OrganizationalMailBox = 0x00000004,
        PrivateDistributionList = 0x00000005,
        RemoteMailUser = 0x00000006,
        Container = 0x00000100,
        Template = 0x00000101,
        OneOffUser = 0x00000102,
        Search = 0x00000200
    }
    #endregion

    /// <summary>
    ///     The PidLidAppointmentUnsendableRecipients  property ([MS-OXPROPS] section 2.35) contains a list of
    ///     unsendable attendees. This property is not required but SHOULD be set
    /// </summary>
    public class UnsendableRecipients : List<RecipientRow>
    {
        #region Properties
        /// <summary>
        ///     An integer that specifies the number of structures in the RecipientRow field
        /// </summary>
        public uint RowCount { get; internal set; }
        #endregion

        #region Constructor
        internal UnsendableRecipients(byte[] data)
        {
            var binaryReader = new BinaryReader(new MemoryStream(data));
            RowCount = binaryReader.ReadUInt32();

            var test = new List<RecipientRow>();

            while (!binaryReader.Eos())
                test.Add(new RecipientRow(binaryReader));
        }
        #endregion
    }

    /// <summary>
    ///     An array of RecipientRow structures, as specified in [MS-OXCDATA] section 2.8.3.
    ///     Each structure specifies an unsendable attendee. The RowCount field specifies the
    ///     number of structures contained in this field. For details about properties that can
    ///     be set on RecipientRow structures for Calendar objects and meeting-related objects,
    ///     see section 2.2.4.10.
    /// </summary>
    /// <remarks>
    ///     See https://msdn.microsoft.com/en-us/library/ee179606(v=exchg.80).aspx
    /// </remarks>
    public class RecipientRow
    {
        #region Properties
        /// <summary>
        ///     The <see cref="RecipientType" />
        /// </summary>
        public RecipientType RecipientType { get; }

        /// <summary>
        ///     The address prefix used
        /// </summary>
        public uint AddressPrefixUsed { get; private set; }

        /// <summary>
        ///     The <see cref="DisplayType" />
        /// </summary>
        public DisplayType DisplayType { get; private set; }

        /// <summary>
        ///     This field MUST be present when the <see cref="RecipientType" /> field of the RecipientFlags
        ///     field is set to X500DN (0x1) and MUST NOT be present otherwise. This value specifies the X500 DN of
        ///     this recipient (1).
        /// </summary>
        /// <remarks>
        ///     A distinguished name (DN), in Teletex form, of an object that is in an address book. An X500 DN can be more limited
        ///     in the size and number of relative distinguished names (RDNs) than a full DN.
        /// </remarks>
        public string X500Dn { get; private set; }

        /// <summary>
        ///     This field MUST be present when the <see cref="RecipientType" /> field of the RecipientFlags field is set to
        ///     PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7). This field MUST
        ///     NOT be present otherwise. This value specifies the size of the EntryID field.
        /// </summary>
        public uint EntryIdSize { get; private set; }

        /// <summary>
        ///     This field MUST be present when the <see cref="RecipientType" /> field of the RecipientFlags field is set to
        ///     PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7). This field MUST NOT be present otherwise. The
        ///     number of bytes in this field MUST be the same as specified in the EntryIdSize field. This array specifies the
        ///     address book EntryID structure, as specified in section 2.2.5.2, of the distribution list.
        /// </summary>
        public AddressBookEntryId EntryId { get; private set; }

        /// <summary>
        ///     This field MUST be present when the <see cref="RecipientType" /> field of the RecipientFlags field is set to
        ///     PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7). This field MUST
        ///     NOT be present otherwise. This value specifies the size of the SearchKey field.
        /// </summary>
        public uint SearchKeySize { get; }

        /// <summary>
        ///     This field is used when the <see cref="RecipientType" /> field of the RecipientFlags field is set to
        ///     PersonalDistributionList1 (0x6) or PersonalDistributionList2 (0x7). This field MUST
        ///     NOT be present otherwise. The number of bytes in this field MUST be the same as what
        ///     is specified in the SearchKeySize field and can be 0. This array specifies the search
        ///     key of the distribution list.
        /// </summary>
        public byte[] SearchKey { get; private set; }

        /// <summary>
        ///     This field MUST be present when the <see cref="RecipientType" /> field of the
        ///     RecipientsFlags field is set to NoType (0x0) and the O flag of the RecipientsFlags field
        ///     is set. This field MUST NOT be present otherwise. This string specifies the address type
        ///     of the recipient (1).
        /// </summary>
        public string AddresType { get; private set; }

        /// <summary>
        ///     A null-terminated string. This field MUST be present when the E flag of the RecipientsFlags
        ///     field is set and MUST NOT be present otherwise. This field MUST be specified in Unicode
        ///     characters if the U flag of the RecipientsFlags field is set and in the 8-bit character set
        ///     otherwise. This string specifies the email address of the recipient (1).
        /// </summary>
        public string EmailAddress { get; private set; }

        /// <summary>
        ///     This field MUST be present when the D flag of the RecipientsFlags
        ///     field is set and MUST NOT be present otherwise. This field MUST be specified in Unicode characters if the U flag of
        ///     the RecipientsFlags field is set and in the 8-bit character set otherwise. This string specifies the email address
        ///     of the recipient (1).
        /// </summary>
        public string DisplayName { get; private set; }

        /// <summary>
        ///     A null-terminated string. This field MUST be present when the I flag of the RecipientsFlags field is set and MUST
        ///     NOT be present otherwise. This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags
        ///     field is set and in the 8-bit character set otherwise. This string specifies the email address of the recipient
        ///     (1).
        /// </summary>
        public string SimpleDisplayName { get; private set; }

        /// <summary>
        ///     This field MUST be present when the T flag of the RecipientsFlags field is set and MUST NOT be present otherwise.
        ///     This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and in the
        ///     8-bit character set otherwise. This string specifies the email address of the recipient (1).
        /// </summary>
        public string TransmittableDisplayName { get; private set; }

        /// <summary>
        ///     PropertyRow structures, as specified in section 2.8.1. The columns used for this row are those specified in
        ///     RecipientProperties.
        /// </summary>
        public List<Property> RecipientProperties { get; private set; }

        /// <summary>
        /// Specifies that the recipient does support receiving rich text messages.
        /// </summary>
        public bool SupportsRtf { get; private set; }
        #endregion

        #region Constructor
        internal RecipientRow(BinaryReader binaryReader)
        {
            // RecipientFlags https://msdn.microsoft.com/en-us/library/ee201786(v=exchg.80).aspx
            var b = new BitArray(binaryReader.ReadBytes(4));

            // If this flag is b'1', a different transport is responsible for delivery to this recipient(1).
            // ReSharper disable once UnusedVariable
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
            SupportsRtf = !b[15];

            switch (RecipientType)
            {
                case RecipientType.X500Dn:
                    AddressPrefixUsed = binaryReader.ReadByte();
                    DisplayType = (DisplayType)binaryReader.ReadByte();
                    X500Dn = Strings.ReadNullTerminatedAsciiString(binaryReader);
                    break;

                case RecipientType.PersonalDistributionList1:
                case RecipientType.PersonalDistributionList2:
                    EntryIdSize = binaryReader.ReadUInt16();
                    EntryId = new AddressBookEntryId(binaryReader);
                    SearchKeySize = binaryReader.ReadUInt16();
                    if (SearchKeySize > 0)
                        SearchKey = binaryReader.ReadBytes((int) SearchKeySize);
                    break;

                case RecipientType.NoType:
                    if (o) AddresType = Strings.ReadNullTerminatedAsciiString(binaryReader);
                    break;
            }

            // MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set
            if (e) EmailAddress = Strings.ReadNullTerminatedString(binaryReader, u);
            if (d) DisplayName = Strings.ReadNullTerminatedString(binaryReader, u);
            if (i) SimpleDisplayName = Strings.ReadNullTerminatedString(binaryReader, u);
            if (s) TransmittableDisplayName = DisplayName;
            else if (t) TransmittableDisplayName = Strings.ReadNullTerminatedString(binaryReader, u);

            // This value specifies the number of columns from the RecipientColumns field that are included in 
            // the RecipientProperties field. 
            var columns = binaryReader.ReadInt16();
            ByteAlign8(binaryReader);

            RecipientProperties = new List<Property>();
            for (var column = 0; column < columns; column++)
            {
                var type = (PropertyType) binaryReader.ReadUInt16();
                var id = binaryReader.ReadUInt16();
                byte[] data;

                switch (type)
                {
                    case PropertyType.PT_NULL:
                    {
                        data = new byte[0];
                        RecipientProperties.Add(new Property(id, type, data));
                        break;
                    }

                    case PropertyType.PT_BOOLEAN:
                    {
                        data = binaryReader.ReadBytes(1);
                        binaryReader.ReadByte();
                        RecipientProperties.Add(new Property(id, type, data));
                        break;
                    }

                    case PropertyType.PT_SHORT:
                    {
                        data = binaryReader.ReadBytes(2);
                        RecipientProperties.Add(new Property(id, type, data));
                        break;
                    }

                    case PropertyType.PT_LONG:
                    case PropertyType.PT_FLOAT:
                    case PropertyType.PT_ERROR:
                    {
                        data = binaryReader.ReadBytes(4);
                        RecipientProperties.Add(new Property(id, type, data));
                        break;
                    }

                    case PropertyType.PT_DOUBLE:
                    case PropertyType.PT_APPTIME:
                    case PropertyType.PT_I8:
                    case PropertyType.PT_SYSTIME:
                    {
                        data = binaryReader.ReadBytes(8);
                        RecipientProperties.Add(new Property(id, type, data));
                        break;
                    }

                    case PropertyType.PT_CLSID:
                    {
                        data = binaryReader.ReadBytes(16);
                        RecipientProperties.Add(new Property(id, type, data));
                        break;
                    }

                    case PropertyType.PT_OBJECT:
                        throw new NotSupportedException("The PT_OBJECT type is not supported");

                    case PropertyType.PT_STRING8:
                    case PropertyType.PT_UNICODE:
                    case PropertyType.PT_BINARY:
                    {
                        var length = binaryReader.ReadInt16();
                        data = binaryReader.ReadBytes(length);
                        RecipientProperties.Add(new Property(id, type, data));
                        break;
                    }

                    case PropertyType.PT_MV_SHORT:
                    {
                        var count = binaryReader.ReadInt16();
                        for (var j = 0; j < count; j++)
                        {
                            data = binaryReader.ReadBytes(2);
                            RecipientProperties.Add(new Property(id, type, data, true));
                        }
                        break;
                    }

                    case PropertyType.PT_MV_LONG:
                    case PropertyType.PT_MV_FLOAT:
                    {
                        var count = binaryReader.ReadInt16();
                        for (var j = 0; j < count; j++)
                        {
                            data = binaryReader.ReadBytes(4);
                            RecipientProperties.Add(new Property(id, type, data, true));
                        }
                        break;
                    }

                    case PropertyType.PT_MV_DOUBLE:
                    case PropertyType.PT_MV_APPTIME:
                    case PropertyType.PT_MV_LONGLONG:
                    case PropertyType.PT_MV_SYSTIME:
                    {
                        var count = binaryReader.ReadInt16();
                        for (var j = 0; j < count; j++)
                        {
                            data = binaryReader.ReadBytes(8);
                            RecipientProperties.Add(new Property(id, type, data, true));
                        }
                        break;
                    }

                    case PropertyType.PT_MV_TSTRING:
                    case PropertyType.PT_MV_STRING8:
                    case PropertyType.PT_MV_BINARY:
                    {
                        var count = binaryReader.ReadInt16();
                        for (var j = 0; j < count; j++)
                        {
                            var length = binaryReader.ReadInt16();
                            data = binaryReader.ReadBytes(length);
                            RecipientProperties.Add(new Property(id, type, data, true));
                        }
                        break;
                    }

                    case PropertyType.PT_MV_CLSID:
                    {
                        var count = binaryReader.ReadInt16();
                        for (var j = 0; j < count; j++)
                        {
                            data = binaryReader.ReadBytes(16);
                            RecipientProperties.Add(new Property(id, type, data));
                        }
                        break;
                    }

                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }
        #endregion

        #region ByteAlign8
        /// <summary>
        /// The dwAlignPad member is used as padding to make sure proper alignment on computers that require 8-byte alignment
        /// for 8-byte values. Developers who write code on such computers should use memory allocation routines that allocate 
        /// the SPropValue arrays on 8-byte boundaries.
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        private static void ByteAlign8(BinaryReader binaryReader)
        {
            var temp = binaryReader.BaseStream.Position;
            while (temp - 8 > 0)
                temp -= 8;

            if (temp != 0)
                binaryReader.BaseStream.Position += (8 - temp);
        }
        #endregion
    }
}