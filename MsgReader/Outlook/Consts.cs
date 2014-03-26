namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal static class Consts
    {
        /// <summary>
        /// (Reserved for interface use) type doesn't matter to caller
        /// </summary>
        public const ushort PtUnspecified = 0;

        /// <summary>
        /// NULL property value
        /// </summary>
        public const ushort PtNull = 1;

        /// <summary>
        /// Signed 16-bit value
        /// </summary>
        public const ushort PtI2 = 2;

        /// <summary>
        /// Signed 32-bit value
        /// </summary>
        public const ushort PtLong = 3;

        /// <summary>
        /// 4-byte floating point
        /// </summary>
        public const ushort PtR4 = 4;

        /// <summary>
        /// Floating point double
        /// </summary>
        public const ushort PtDouble = 5;

        /// <summary>
        /// Signed 64-bit int (decimal w/4 digits right of decimal pt)
        /// </summary>
        public const ushort PtCurrency = 6;

        /// <summary>
        /// Application time
        /// </summary>
        public const ushort PtApptime = 7;

        /// <summary>
        /// 32-bit error value
        /// </summary>
        public const ushort PtError = 10;

        /// <summary>
        /// 16-bit boolean (non-zero true)
        /// </summary>
        public const ushort PtBoolean = 11;

        /// <summary>
        /// Embedded object in a property
        /// </summary>
        public const ushort PtObject = 13;

        /// <summary>
        /// 8-byte signed integer
        /// </summary>
        public const ushort PtI8 = 20;

        /// <summary>
        /// Null terminated 8-bit character string
        /// </summary>
        public const ushort PtString8 = 30;

        /// <summary>
        /// Null terminated Unicode string
        /// </summary>
        public const ushort PtUnicode = 31;

        /// <summary>
        /// FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601
        /// </summary>
        public const ushort PtSystime = 64;

        /// <summary>
        /// OLE GUID
        /// </summary>
        public const ushort PtClsid = 72;

        /// <summary>
        /// Uninterpreted (counted byte array)
        /// </summary>
        public const ushort PtBinary = 258;

        /// <summary>
        /// Multi-view unicode string
        /// </summary>
        public const ushort PtMvUnicode = 4127;

        /// <summary>
        /// Gives the type of class that is used for the msg file:
        /// - IPM.Note = E-mail
        /// - IMP.Appointment = "Agenda item" 
        /// </summary>
        public const string PrMessageClass = "001A";

        /// <summary>
        /// Prefix that is placed before an attachment tag
        /// </summary>
        public const string AttachStoragePrefix = "__attach_version1.0_#";

        /// <summary>
        /// Name of the attachment
        /// </summary>
        public const string PrAttachFilename = "3704";

        /// <summary>
        /// Long name of the attachment
        /// </summary>
        public const string PrAttachLongFilename = "3707";

        /// <summary>
        /// Attachment data
        /// </summary>
        public const string PrAttachData = "3701";

        /// <summary>
        /// Method that is used to embed the attachment e.g. inline
        /// </summary>
        public const string PrAttachMethod = "3705";

        /// <summary>
        /// Position in the HTML body for the attachment
        /// </summary>
        public const string PrRenderingPosition = "370B";

        /// <summary>
        /// Content ID from the attachment
        /// </summary>
        public const string PrAttachContentId = "3712";

        public const int AttachEmbeddedMsg = 5;
        public const string RecipStoragePrefix = "__recip_version1.0_#";

        /// <summary>
        /// Displayname e.g. Kees van Spelde
        /// </summary>
        public const string PrDisplayName = "3001";

        /// <summary>
        /// E-mail address e.g. PeterPan@neverland.com
        /// </summary>
        public const string PrEmail = "39FE";
        public const string PrEmail2 = "403E";
        public const string PrRecipientType = "0C15";

        /// <summary>
        /// E-mail To
        /// </summary>
        public const int MapiTo = 1;

        /// <summary>
        /// E-mail From
        /// </summary>
        public const int MapiCc = 2;

        /// <summary>
        /// E-mail subject
        /// </summary>
        public const string PrSubject = "0037";

        /// <summary>
        /// The date and time when the E-mail has been sent
        /// </summary>
        public const string PrClientSubmitTime = "0039";

        /// <summary>
        /// The date and time when the E-mail has been delivered to the addressee
        /// </summary>
        public const string PrMessageDeliveryTime = "0E06";

        /// <summary>
        /// The text body of the E-mail
        /// </summary>
        public const string PrBody = "1000";

        /// <summary>
        /// The HTML body of the E-mail
        /// </summary>
        public const string PrBodyHtml = "1013";

        /// <summary>
        /// The RTF body of the E-mail, this one can contain the HTML body embedded into the RTF
        /// </summary>
        public const string PrRtfCompressed = "1009";

        /// <summary>
        /// Gives the codepage used in the body or html
        /// </summary>
        public const string PrInternetCpid = "3FDE"; 

        /// <summary>
        /// Name of the sender e.g. Kees van Spelde
        /// </summary>
        public const string PrSenderName = "0C1A";

        /// <summary>
        /// E-mail address of the sender e.g. PeterPan@neverland.com
        /// </summary>
        public const string PrSenderEmail = "0C1F";

        /// <summary>
        /// E-mail address of the sender e.g. PeterPan@neverland.com
        /// </summary>
        public const string PrSenderEmail2 = "8012";

        /// <summary>
        /// Specifies the color to be used when displaying a Calendar object
        /// </summary>
        public const string PidNameKeywords = "80A4";

        /// <summary>
        /// Can contain the internet E-mail headers
        /// </summary>
        public const string PrTransportMessageHeaders1 = "007D001E";

        /// <summary>
        /// Can contain the internet E-mail headers
        /// </summary>
        public const string PrTransportMessageHeaders2 = "007D001F";
        public const string HeaderStreamName = "__substg1.0_007D001F";
        public const string PropertiesStream = "__properties_version1.0";
        public const int PropertiesStreamHeaderTop = 32;
        public const int PropertiesStreamHeaderEmbeded = 24;
        public const int PropertiesStreamHeaderAttachOrRecip = 8;
        public const string NameidStorage = "__nameid_version1.0";
    }
}
