namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    /// <summary>
    ///     Contains all MAPI related constants
    /// </summary>
    internal static class MapiTags
    {
        // ReSharper disable InconsistentNaming
        /*
         *	M A P I T A G S . H
         *
         *	Property tag definitions for standard properties of MAPI
         *	objects.
         *
         *	The following ranges should be used for all property IDs. Note that
         *	property IDs for objects other than messages and recipients should
         *	all fall in the range "3000 to "3FFF:
         *
         *	From	To		Kind of property
         *	--------------------------------
         *	0001	0BFF	MAPI_defined envelope property
         *	0C00	0DFF	MAPI_defined per-recipient property
         *	0E00	0FFF	MAPI_defined non-transmittable property
         *	1000	2FFF	MAPI_defined message content property
         *
         *	3000	3FFF	MAPI_defined property (usually not message or recipient)
         *
         *	4000	57FF	Transport-defined envelope property
         *	5800	5FFF	Transport-defined per-recipient property
         *	6000	65FF	User-defined non-transmittable property
         *	6600	67FF	Provider-defined internal non-transmittable property
         *	6800	7BFF	Message class-defined content property
         *	7C00	7FFF	Message class-defined non-transmittable
         *					property
         *
         *	8000	FFFE	User-defined Name-to-id mapped property
         *
         *	The 3000-3FFF range is further subdivided as follows:
         *
         *	From	To		Kind of property
         *	--------------------------------
         *	3000	33FF	Common property such as display name, entry ID
         *	3400	35FF	Message store object
         *	3600	36FF	Folder or AB container
         *	3700	38FF	Attachment
         *	3900	39FF	Address book object
         *	3A00	3BFF	Mail user
         *	3C00	3CFF	Distribution list
         *	3D00	3DFF	Profile section
         *	3E00	3FFF	Status object
         *
         *  Copyright (c" 2009 Microsoft Corporation. All Rights Reserved.
         */

        #region Mapi standard tags
        public const string PR_ACKNOWLEDGEMENT_MODE = "0001";
        public const string PR_ALTERNATE_RECIPIENT_ALLOWED = "0002";
        public const string PR_AUTHORIZING_USERS = "0003";
        public const string PR_AUTO_FORWARD_COMMENT = "0004";
        public const string PR_AUTO_FORWARD_COMMENT_W = "0004";
        public const string PR_AUTO_FORWARD_COMMENT_A = "0004";
        public const string PR_AUTO_FORWARDED = "0005";
        public const string PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID = "0006";
        public const string PR_CONTENT_CORRELATOR = "0007";
        public const string PR_CONTENT_IDENTIFIER = "0008";
        public const string PR_CONTENT_IDENTIFIER_W = "0008";
        public const string PR_CONTENT_IDENTIFIER_A = "0008";
        public const string PR_CONTENT_LENGTH = "0009";
        public const string PR_CONTENT_RETURN_REQUESTED = "000A";
        public const string PR_CONVERSATION_KEY = "000B";
        public const string PR_CONVERSION_EITS = "000C";
        public const string PR_CONVERSION_WITH_LOSS_PROHIBITED = "000D";
        public const string PR_CONVERTED_EITS = "000E";
        public const string PR_DEFERRED_DELIVERY_TIME = "000F";
        public const string PR_DELIVER_TIME = "0010";
        public const string PR_DISCARD_REASON = "0011";
        public const string PR_DISCLOSURE_OF_RECIPIENTS = "0012";
        public const string PR_DL_EXPANSION_HISTORY = "0013";
        public const string PR_DL_EXPANSION_PROHIBITED = "0014";
        public const string PR_EXPIRY_TIME = "0015";
        public const string PR_IMPLICIT_CONVERSION_PROHIBITED = "0016";
        public const string PR_IMPORTANCE = "0017";
        public const string PR_IPM_ID = "0018";
        public const string PR_LATEST_DELIVERY_TIME = "0019";
        public const string PR_MESSAGE_CLASS = "001A";
        public const string PR_MESSAGE_CLASS_W = "001A";
        public const string PR_MESSAGE_CLASS_A = "001A";
        public const string PR_MESSAGE_DELIVERY_ID = "001B";
        public const string PR_MESSAGE_SECURITY_LABEL = "001E";
        public const string PR_OBSOLETED_IPMS = "001F";
        public const string PR_ORIGINALLY_INTENDED_RECIPIENT_NAME = "0020";
        public const string PR_ORIGINAL_EITS = "0021";
        public const string PR_ORIGINATOR_CERTIFICATE = "0022";
        public const string PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED = "0023";
        public const string PR_ORIGINATOR_RETURN_ADDRESS = "0024";
        public const string PR_PARENT_KEY = "0025";
        public const string PR_PRIORITY = "0026";
        public const string PR_ORIGIN_CHECK = "0027";
        public const string PR_PROOF_OF_SUBMISSION_REQUESTED = "0028";
        public const string PR_READ_RECEIPT_REQUESTED = "0029";
        public const string PR_RECEIPT_TIME = "002A";
        public const string PR_RECIPIENT_REASSIGNMENT_PROHIBITED = "002B";
        public const string PR_REDIRECTION_HISTORY = "002C";
        public const string PR_RELATED_IPMS = "002D";
        public const string PR_ORIGINAL_SENSITIVITY = "002E";
        public const string PR_LANGUAGES = "002F";
        public const string PR_LANGUAGES_W = "002F";
        public const string PR_LANGUAGES_A = "002F";
        public const string PR_REPLY_TIME = "0030";
        public const string PR_REPORT_TAG = "0031";
        public const string PR_REPORT_TIME = "0032";
        public const string PR_RETURNED_IPM = "0033";
        public const string PR_SECURITY = "0034";
        public const string PR_INCOMPLETE_COPY = "0035";
        public const string PR_SENSITIVITY = "0036";
        public const string PR_SUBJECT = "0037";
        public const string PR_SUBJECT_W = "0037";
        public const string PR_SUBJECT_A = "0037";
        public const string PR_SUBJECT_IPM = "0038";
        public const string PR_CLIENT_SUBMIT_TIME = "0039";
        public const string PR_REPORT_NAME = "003A";
        public const string PR_REPORT_NAME_W = "003A";
        public const string PR_REPORT_NAME_A = "003A";
        public const string PR_SENT_REPRESENTING_SEARCH_KEY = "003B";
        public const string PR_X400_CONTENT_TYPE = "003C";
        public const string PR_SUBJECT_PREFIX = "003D";
        public const string PR_SUBJECT_PREFIX_W = "003D";
        public const string PR_SUBJECT_PREFIX_A = "003D";
        public const string PR_NON_RECEIPT_REASON = "003E";
        public const string PR_RECEIVED_BY_ENTRYID = "003F";
        public const string PR_RECEIVED_BY_NAME = "0040";
        public const string PR_RECEIVED_BY_NAME_W = "0040";
        public const string PR_RECEIVED_BY_NAME_A = "0040";
        public const string PR_SENT_REPRESENTING_ENTRYID = "0041";
        public const string PR_SENT_REPRESENTING_NAME = "0042";
        public const string PR_SENT_REPRESENTING_NAME_W = "0042";
        public const string PR_SENT_REPRESENTING_NAME_A = "0042";
        public const string PR_RCVD_REPRESENTING_ENTRYID = "0043";
        public const string PR_RCVD_REPRESENTING_NAME = "0044";
        public const string PR_RCVD_REPRESENTING_NAME_W = "0044";
        public const string PR_RCVD_REPRESENTING_NAME_A = "0044";
        public const string PR_REPORT_ENTRYID = "0045";
        public const string PR_READ_RECEIPT_ENTRYID = "0046";
        public const string PR_MESSAGE_SUBMISSION_ID = "0047";
        public const string PR_PROVIDER_SUBMIT_TIME = "0048";
        public const string PR_ORIGINAL_SUBJECT = "0049";
        public const string PR_ORIGINAL_SUBJECT_W = "0049";
        public const string PR_ORIGINAL_SUBJECT_A = "0049";
        public const string PR_DISC_VAL = "004A";
        public const string PR_ORIG_MESSAGE_CLASS = "004B";
        public const string PR_ORIG_MESSAGE_CLASS_W = "004B";
        public const string PR_ORIG_MESSAGE_CLASS_A = "004B";
        public const string PR_ORIGINAL_AUTHOR_ENTRYID = "004C";
        public const string PR_ORIGINAL_AUTHOR_NAME = "004D";
        public const string PR_ORIGINAL_AUTHOR_NAME_W = "004D";
        public const string PR_ORIGINAL_AUTHOR_NAME_A = "004D";
        public const string PR_ORIGINAL_SUBMIT_TIME = "004E";
        public const string PR_REPLY_RECIPIENT_ENTRIES = "004F";
        public const string PR_REPLY_RECIPIENT_NAMES = "0050";
        public const string PR_REPLY_RECIPIENT_NAMES_W = "0050";
        public const string PR_REPLY_RECIPIENT_NAMES_A = "0050";

        public const string PR_RECEIVED_BY_SEARCH_KEY = "0051";
        public const string PR_RCVD_REPRESENTING_SEARCH_KEY = "0052";
        public const string PR_READ_RECEIPT_SEARCH_KEY = "0053";
        public const string PR_REPORT_SEARCH_KEY = "0054";
        public const string PR_ORIGINAL_DELIVERY_TIME = "0055";
        public const string PR_ORIGINAL_AUTHOR_SEARCH_KEY = "0056";

        public const string PR_MESSAGE_TO_ME = "0057";
        public const string PR_MESSAGE_CC_ME = "0058";
        public const string PR_MESSAGE_RECIP_ME = "0059";

        public const string PR_ORIGINAL_SENDER_NAME = "005A";
        public const string PR_ORIGINAL_SENDER_NAME_W = "005A";
        public const string PR_ORIGINAL_SENDER_NAME_A = "005A";
        public const string PR_ORIGINAL_SENDER_ENTRYID = "005B";
        public const string PR_ORIGINAL_SENDER_SEARCH_KEY = "005C";
        public const string PR_ORIGINAL_SENT_REPRESENTING_NAME = "005D";
        public const string PR_ORIGINAL_SENT_REPRESENTING_NAME_W = "005D";
        public const string PR_ORIGINAL_SENT_REPRESENTING_NAME_A = "005D";
        public const string PR_ORIGINAL_SENT_REPRESENTING_ENTRYID = "005E";
        public const string PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY = "005F";

        public const string PR_START_DATE = "0060";
        public const string PR_END_DATE = "0061";
        public const string PR_OWNER_APPT_ID = "0062";
        public const string PR_RESPONSE_REQUESTED = "0063";

        public const string PR_SENT_REPRESENTING_ADDRTYPE = "0064";
        public const string PR_SENT_REPRESENTING_ADDRTYPE_W = "0064";
        public const string PR_SENT_REPRESENTING_ADDRTYPE_A = "0064";
        public const string PR_SENT_REPRESENTING_EMAIL_ADDRESS = "0065";
        public const string PR_SENT_REPRESENTING_EMAIL_ADDRESS_W = "0065";
        public const string PR_SENT_REPRESENTING_EMAIL_ADDRESS_A = "0065";

        public const string PR_ORIGINAL_SENDER_ADDRTYPE = "0066";
        public const string PR_ORIGINAL_SENDER_ADDRTYPE_W = "0066";
        public const string PR_ORIGINAL_SENDER_ADDRTYPE_A = "0066";
        public const string PR_ORIGINAL_SENDER_EMAIL_ADDRESS = "0067";
        public const string PR_ORIGINAL_SENDER_EMAIL_ADDRESS_W = "0067";
        public const string PR_ORIGINAL_SENDER_EMAIL_ADDRESS_A = "0067";

        public const string PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE = "0068";
        public const string PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_W = "0068";
        public const string PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_A = "0068";
        public const string PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS = "0069";
        public const string PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_W = "0069";
        public const string PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_A = "0069";

        public const string PR_CONVERSATION_TOPIC = "0070";
        public const string PR_CONVERSATION_TOPIC_W = "0070";
        public const string PR_CONVERSATION_TOPIC_A = "0070";
        public const string PR_CONVERSATION_INDEX = "0071";

        public const string PR_ORIGINAL_DISPLAY_BCC = "0072";
        public const string PR_ORIGINAL_DISPLAY_BCC_W = "0072";
        public const string PR_ORIGINAL_DISPLAY_BCC_A = "0072";
        public const string PR_ORIGINAL_DISPLAY_CC = "0073";
        public const string PR_ORIGINAL_DISPLAY_CC_W = "0073";
        public const string PR_ORIGINAL_DISPLAY_CC_A = "0073";
        public const string PR_ORIGINAL_DISPLAY_TO = "0074";
        public const string PR_ORIGINAL_DISPLAY_TO_W = "0074";
        public const string PR_ORIGINAL_DISPLAY_TO_A = "0074";

        public const string PR_RECEIVED_BY_ADDRTYPE = "0075";
        public const string PR_RECEIVED_BY_ADDRTYPE_W = "0075";
        public const string PR_RECEIVED_BY_ADDRTYPE_A = "0075";
        public const string PR_RECEIVED_BY_EMAIL_ADDRESS = "0076";
        public const string PR_RECEIVED_BY_EMAIL_ADDRESS_W = "0076";
        public const string PR_RECEIVED_BY_EMAIL_ADDRESS_A = "0076";

        public const string PR_RCVD_REPRESENTING_ADDRTYPE = "0077";
        public const string PR_RCVD_REPRESENTING_ADDRTYPE_W = "0077";
        public const string PR_RCVD_REPRESENTING_ADDRTYPE_A = "0077";
        public const string PR_RCVD_REPRESENTING_EMAIL_ADDRESS = "0078";
        public const string PR_RCVD_REPRESENTING_EMAIL_ADDRESS_W = "0078";
        public const string PR_RCVD_REPRESENTING_EMAIL_ADDRESS_A = "0078";

        public const string PR_ORIGINAL_AUTHOR_ADDRTYPE = "0079";
        public const string PR_ORIGINAL_AUTHOR_ADDRTYPE_W = "0079";
        public const string PR_ORIGINAL_AUTHOR_ADDRTYPE_A = "0079";
        public const string PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS = "007A";
        public const string PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_W = "007A";
        public const string PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_A = "007A";

        public const string PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE = "007B";
        public const string PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_W = "007B";
        public const string PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_A = "007B";
        public const string PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS = "007C";
        public const string PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_W = "007C";
        public const string PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_A = "007C";

        public const string PR_TRANSPORT_MESSAGE_HEADERS = "007D";
        public const string PR_TRANSPORT_MESSAGE_HEADERS_W = "007D";
        public const string PR_TRANSPORT_MESSAGE_HEADERS_A = "007D";
        public const string PR_DELEGATION = "007E";
        public const string PR_TNEF_CORRELATION_KEY = "007F";

        /*
         *	Message content properties
         */

        public const string PR_BODY = "1000";
        public const string PR_BODY_W = "1000";
        public const string PR_BODY_A = "1000";
        public const string PR_BODY_HTML = "1013";
        public const string PR_REPORT_TEXT = "1001";
        public const string PR_REPORT_TEXT_W = "1001";
        public const string PR_REPORT_TEXT_A = "1001";
        public const string PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY = "1002";
        public const string PR_REPORTING_DL_NAME = "1003";
        public const string PR_REPORTING_MTA_CERTIFICATE = "1004";

        /*  Removed PR_REPORT_ORIGIN_AUTHENTICATION_CHECK with DCR 3865, use PR_ORIGIN_CHECK */

        public const string PR_RTF_SYNC_BODY_CRC = "1006";
        public const string PR_RTF_SYNC_BODY_COUNT = "1007";
        public const string PR_RTF_SYNC_BODY_TAG = "1008";
        public const string PR_RTF_SYNC_BODY_TAG_W = "1008";
        public const string PR_RTF_SYNC_BODY_TAG_A = "1008";
        public const string PR_RTF_COMPRESSED = "1009";
        public const string PR_RTF_SYNC_PREFIX_COUNT = "1010";
        public const string PR_RTF_SYNC_TRAILING_COUNT = "1011";
        public const string PR_ORIGINALLY_INTENDED_RECIP_ENTRYID = "1012";
        
        /*
         *  Reserved "1100-"1200
         */


        /*
         *	Message recipient properties
         */

        public const string PR_CONTENT_INTEGRITY_CHECK = "0C00";
        public const string PR_EXPLICIT_CONVERSION = "0C01";
        public const string PR_IPM_RETURN_REQUESTED = "0C02";
        public const string PR_MESSAGE_TOKEN = "0C03";
        public const string PR_NDR_REASON_CODE = "0C04";
        public const string PR_NDR_DIAG_CODE = "0C05";
        public const string PR_NON_RECEIPT_NOTIFICATION_REQUESTED = "0C06";
        public const string PR_DELIVERY_POINT = "0C07";

        public const string PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED = "0C08";
        public const string PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT = "0C09";
        public const string PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY = "0C0A";
        public const string PR_PHYSICAL_DELIVERY_MODE = "0C0B";
        public const string PR_PHYSICAL_DELIVERY_REPORT_REQUEST = "0C0C";
        public const string PR_PHYSICAL_FORWARDING_ADDRESS = "0C0D";
        public const string PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED = "0C0E";
        public const string PR_PHYSICAL_FORWARDING_PROHIBITED = "0C0F";
        public const string PR_PHYSICAL_RENDITION_ATTRIBUTES = "0C10";
        public const string PR_PROOF_OF_DELIVERY = "0C11";
        public const string PR_PROOF_OF_DELIVERY_REQUESTED = "0C12";
        public const string PR_RECIPIENT_CERTIFICATE = "0C13";
        public const string PR_RECIPIENT_NUMBER_FOR_ADVICE = "0C14";
        public const string PR_RECIPIENT_NUMBER_FOR_ADVICE_W = "0C14";
        public const string PR_RECIPIENT_NUMBER_FOR_ADVICE_A = "0C14";
        public const string PR_RECIPIENT_TYPE = "0C15";
        public const string PR_REGISTERED_MAIL_TYPE = "0C16";
        public const string PR_REPLY_REQUESTED = "0C17";
        public const string PR_REQUESTED_DELIVERY_METHOD = "0C18";
        public const string PR_SENDER_ENTRYID = "0C19";
        public const string PR_SENDER_NAME = "0C1A";
        public const string PR_SENDER_NAME_W = "0C1A";
        public const string PR_SENDER_NAME_A = "0C1A";
        public const string PR_SUPPLEMENTARY_INFO = "0C1B";
        public const string PR_SUPPLEMENTARY_INFO_W = "0C1B";
        public const string PR_SUPPLEMENTARY_INFO_A = "0C1B";
        public const string PR_TYPE_OF_MTS_USER = "0C1C";
        public const string PR_SENDER_SEARCH_KEY = "0C1D";
        public const string PR_SENDER_ADDRTYPE = "0C1E";
        public const string PR_SENDER_ADDRTYPE_W = "0C1E";
        public const string PR_SENDER_ADDRTYPE_A = "0C1E";
        public const string PR_SENDER_EMAIL_ADDRESS = "0C1F";
        public const string PR_SENDER_EMAIL_ADDRESS_W = "0C1F";
        public const string PR_SENDER_EMAIL_ADDRESS_A = "0C1F";

        /*
         *	Message non-transmittable properties
         */

        /*
         * The two tags, PR_MESSAGE_RECIPIENTS and PR_MESSAGE_ATTACHMENTS,
         * are to be used in the exclude list passed to
         * IMessage::CopyTo when the caller wants either the recipients or attachments
         * of the message to not get copied.  It is also used in the ProblemArray
         * return from IMessage::CopyTo when an error is encountered copying them
         */

        public const string PR_CURRENT_VERSION = "0E00";
        public const string PR_DELETE_AFTER_SUBMIT = "0E01";
        public const string PR_DISPLAY_BCC = "0E02";
        public const string PR_DISPLAY_BCC_W = "0E02";
        public const string PR_DISPLAY_BCC_A = "0E02";
        public const string PR_DISPLAY_CC = "0E03";
        public const string PR_DISPLAY_CC_W = "0E03";
        public const string PR_DISPLAY_CC_A = "0E03";
        public const string PR_DISPLAY_TO = "0E04";
        public const string PR_DISPLAY_TO_W = "0E04";
        public const string PR_DISPLAY_TO_A = "0E04";
        public const string PR_PARENT_DISPLAY = "0E05";
        public const string PR_PARENT_DISPLAY_W = "0E05";
        public const string PR_PARENT_DISPLAY_A = "0E05";
        public const string PR_MESSAGE_DELIVERY_TIME = "0E06";
        public const string PR_MESSAGE_FLAGS = "0E07";
        public const string PR_MESSAGE_SIZE = "0E08";
        public const string PR_PARENT_ENTRYID = "0E09";
        public const string PR_SENTMAIL_ENTRYID = "0E0A";
        public const string PR_CORRELATE = "0E0C";
        public const string PR_CORRELATE_MTSID = "0E0D";
        public const string PR_DISCRETE_VALUES = "0E0E";
        public const string PR_RESPONSIBILITY = "0E0F";
        public const string PR_SPOOLER_STATUS = "0E10";
        public const string PR_TRANSPORT_STATUS = "0E11";
        public const string PR_MESSAGE_RECIPIENTS = "0E12";
        public const string PR_MESSAGE_ATTACHMENTS = "0E13";
        public const string PR_SUBMIT_FLAGS = "0E14";
        public const string PR_RECIPIENT_STATUS = "0E15";
        public const string PR_TRANSPORT_KEY = "0E16";
        public const string PR_MSG_STATUS = "0E17";
        public const string PR_MESSAGE_DOWNLOAD_TIME = "0E18";
        public const string PR_CREATION_VERSION = "0E19";
        public const string PR_MODIFY_VERSION = "0E1A";
        public const string PR_HASATTACH = "0E1B";
        public const string PR_BODY_CRC = "0E1C";
        public const string PR_NORMALIZED_SUBJECT = "0E1D";
        public const string PR_NORMALIZED_SUBJECT_W = "0E1D";
        public const string PR_NORMALIZED_SUBJECT_A = "0E1D";
        public const string PR_RTF_IN_SYNC = "0E1F";
        public const string PR_ATTACH_SIZE = "0E20";
        public const string PR_ATTACH_NUM = "0E21";
        public const string PR_PREPROCESS = "0E22";

        /* PR_ORIGINAL_DISPLAY_TO, _CC, and _BCC moved to transmittible range 03/09/95 */

        public const string PR_ORIGINATING_MTA_CERTIFICATE = "0E25";
        public const string PR_PROOF_OF_SUBMISSION = "0E26";

        /*
         * The range of non-message and non-recipient property IDs ("3000 - "3FFF" is
         * further broken down into ranges to make assigning new property IDs easier.
         *
         *	From	To		Kind of property
         *	--------------------------------
         *	3000	32FF	MAPI_defined common property
         *	3200	33FF	MAPI_defined form property
         *	3400	35FF	MAPI_defined message store property
         *	3600	36FF	MAPI_defined Folder or AB Container property
         *	3700	38FF	MAPI_defined attachment property
         *	3900	39FF	MAPI_defined address book property
         *	3A00	3BFF	MAPI_defined mailuser property
         *	3C00	3CFF	MAPI_defined DistList property
         *	3D00	3DFF	MAPI_defined Profile Section property
         *	3E00	3EFF	MAPI_defined Status property
         *	3F00	3FFF	MAPI_defined display table property
         */

        /*
         *	Properties common to numerous MAPI objects.
         *
         *	Those properties that can appear on messages are in the
         *	non-transmittable range for messages. They start at the high
         *	end of that range and work down.
         *
         *	Properties that never appear on messages are defined in the common
         *	property range (see above".
         */

        /*
         * properties that are common to multiple objects (including message objects";
         * -- these ids are in the non-transmittable range
         */

        public const string PR_ENTRYID = "0FFF";
        public const string PR_OBJECT_TYPE = "0FFE";
        public const string PR_ICON = "0FFD";
        public const string PR_MINI_ICON = "0FFC";
        public const string PR_STORE_ENTRYID = "0FFB";
        public const string PR_STORE_RECORD_KEY = "0FFA";
        public const string PR_RECORD_KEY = "0FF9";
        public const string PR_MAPPING_SIGNATURE = "0FF8";
        public const string PR_ACCESS_LEVEL = "0FF7";
        public const string PR_INSTANCE_KEY = "0FF6";
        public const string PR_ROW_TYPE = "0FF5";
        public const string PR_ACCESS = "0FF4";

        /*
         * properties that are common to multiple objects (usually not including message objects";
         * -- these ids are in the transmittable range
         */

        public const string PR_ROWID = "3000";
        public const string PR_DISPLAY_NAME = "3001";
        public const string PR_DISPLAY_NAME_W = "3001";
        public const string PR_DISPLAY_NAME_A = "3001";
        public const string PR_ADDRTYPE = "3002";
        public const string PR_ADDRTYPE_W = "3002";
        public const string PR_ADDRTYPE_A = "3002";
        public const string PR_EMAIL_ADDRESS = "3003";
        public const string PR_EMAIL_ADDRESS_W = "3003";
        public const string PR_EMAIL_ADDRESS_A = "3003";
        public const string PR_COMMENT = "3004";
        public const string PR_COMMENT_W = "3004";
        public const string PR_COMMENT_A = "3004";
        public const string PR_DEPTH = "3005";
        public const string PR_PROVIDER_DISPLAY = "3006";
        public const string PR_PROVIDER_DISPLAY_W = "3006";
        public const string PR_PROVIDER_DISPLAY_A = "3006";
        public const string PR_CREATION_TIME = "3007";
        public const string PR_LAST_MODIFICATION_TIME = "3008";
        public const string PR_RESOURCE_FLAGS = "3009";
        public const string PR_PROVIDER_DLL_NAME = "300A";
        public const string PR_PROVIDER_DLL_NAME_W = "300A";
        public const string PR_PROVIDER_DLL_NAME_A = "300A";
        public const string PR_SEARCH_KEY = "300B";
        public const string PR_PROVIDER_UID = "300C";
        public const string PR_PROVIDER_ORDINAL = "300D";

        /*
         *  MAPI Form properties
         */
        public const string PR_FORM_VERSION = "3301";
        public const string PR_FORM_VERSION_W = "3301";
        public const string PR_FORM_VERSION_A = "3301";
        public const string PR_FORM_CLSID = "3302";
        public const string PR_FORM_CONTACT_NAME = "3303";
        public const string PR_FORM_CONTACT_NAME_W = "3303";
        public const string PR_FORM_CONTACT_NAME_A = "3303";
        public const string PR_FORM_CATEGORY = "3304";
        public const string PR_FORM_CATEGORY_W = "3304";
        public const string PR_FORM_CATEGORY_A = "3304";
        public const string PR_FORM_CATEGORY_SUB = "3305";
        public const string PR_FORM_CATEGORY_SUB_W = "3305";
        public const string PR_FORM_CATEGORY_SUB_A = "3305";
        public const string PR_FORM_HOST_MAP = "3306";
        public const string PR_FORM_HIDDEN = "3307";
        public const string PR_FORM_DESIGNER_NAME = "3308";
        public const string PR_FORM_DESIGNER_NAME_W = "3308";
        public const string PR_FORM_DESIGNER_NAME_A = "3308";
        public const string PR_FORM_DESIGNER_GUID = "3309";
        public const string PR_FORM_MESSAGE_BEHAVIOR = "330A";

        /*
         *	Message store properties
         */

        public const string PR_DEFAULT_STORE = "3400";
        public const string PR_STORE_SUPPORT_MASK = "340D";
        public const string PR_STORE_STATE = "340E";

        public const string PR_IPM_SUBTREE_SEARCH_KEY = "3410";
        public const string PR_IPM_OUTBOX_SEARCH_KEY = "3411";
        public const string PR_IPM_WASTEBASKET_SEARCH_KEY = "3412";
        public const string PR_IPM_SENTMAIL_SEARCH_KEY = "3413";
        public const string PR_MDB_PROVIDER = "3414";
        public const string PR_RECEIVE_FOLDER_SETTINGS = "3415";

        public const string PR_VALID_FOLDER_MASK = "35DF";
        public const string PR_IPM_SUBTREE_ENTRYID = "35E0";

        public const string PR_IPM_OUTBOX_ENTRYID = "35E2";
        public const string PR_IPM_WASTEBASKET_ENTRYID = "35E3";
        public const string PR_IPM_SENTMAIL_ENTRYID = "35E4";
        public const string PR_VIEWS_ENTRYID = "35E5";
        public const string PR_COMMON_VIEWS_ENTRYID = "35E6";
        public const string PR_FINDER_ENTRYID = "35E7";

        /* Proptags "35E8-"35FF reserved for folders "guaranteed" by PR_VALID_FOLDER_MASK */


        /*
         *	Folder and AB Container properties
         */

        public const string PR_CONTAINER_FLAGS = "3600";
        public const string PR_FOLDER_TYPE = "3601";
        public const string PR_CONTENT_COUNT = "3602";
        public const string PR_CONTENT_UNREAD = "3603";
        public const string PR_CREATE_TEMPLATES = "3604";
        public const string PR_DETAILS_TABLE = "3605";
        public const string PR_SEARCH = "3607";
        public const string PR_SELECTABLE = "3609";
        public const string PR_SUBFOLDERS = "360A";
        public const string PR_STATUS = "360B";
        public const string PR_ANR = "360C";
        public const string PR_ANR_W = "360C";
        public const string PR_ANR_A = "360C";
        public const string PR_CONTENTS_SORT_ORDER = "360D";
        public const string PR_CONTAINER_HIERARCHY = "360E";
        public const string PR_CONTAINER_CONTENTS = "360F";
        public const string PR_FOLDER_ASSOCIATED_CONTENTS = "3610";
        public const string PR_DEF_CREATE_DL = "3611";
        public const string PR_DEF_CREATE_MAILUSER = "3612";
        public const string PR_CONTAINER_CLASS = "3613";
        public const string PR_CONTAINER_CLASS_W = "3613";
        public const string PR_CONTAINER_CLASS_A = "3613";
        public const string PR_CONTAINER_MODIFY_VERSION = "3614";
        public const string PR_AB_PROVIDER_ID = "3615";
        public const string PR_DEFAULT_VIEW_ENTRYID = "3616";
        public const string PR_ASSOC_CONTENT_COUNT = "3617";

        /* Reserved "36C0-"36FF */

        /*
         *	Attachment properties
         */

        public const string PR_ATTACHMENT_X400_PARAMETERS = "3700";
        public const string PR_ATTACH_DATA_OBJ = "3701";
        public const string PR_ATTACH_DATA_BIN = "3701";
        public const string PR_ATTACH_ENCODING = "3702";
        public const string PR_ATTACH_EXTENSION = "3703";
        public const string PR_ATTACH_EXTENSION_W = "3703";
        public const string PR_ATTACH_EXTENSION_A = "3703";
        public const string PR_ATTACH_FILENAME = "3704";
        public const string PR_ATTACH_FILENAME_W = "3704";
        public const string PR_ATTACH_FILENAME_A = "3704";
        public const string PR_ATTACH_METHOD = "3705";
        public const string PR_ATTACH_LONG_FILENAME = "3707";
        public const string PR_ATTACH_LONG_FILENAME_W = "3707";
        public const string PR_ATTACH_LONG_FILENAME_A = "3707";
        public const string PR_ATTACH_PATHNAME = "3708";
        public const string PR_ATTACH_PATHNAME_W = "3708";
        public const string PR_ATTACH_PATHNAME_A = "3708";
        public const string PR_ATTACH_RENDERING = "3709";
        public const string PR_ATTACH_CONTENTID = "3712";
        public const string PR_ATTACH_TAG = "370A";
        public const string PR_RENDERING_POSITION = "370B";
        public const string PR_ATTACH_TRANSPORT_NAME = "370C";
        public const string PR_ATTACH_TRANSPORT_NAME_W = "370C";
        public const string PR_ATTACH_TRANSPORT_NAME_A = "370C";
        public const string PR_ATTACH_LONG_PATHNAME = "370D";
        public const string PR_ATTACH_LONG_PATHNAME_W = "370D";
        public const string PR_ATTACH_LONG_PATHNAME_A = "370D";
        public const string PR_ATTACH_MIME_TAG = "370E";
        public const string PR_ATTACH_MIME_TAG_W = "370E";
        public const string PR_ATTACH_MIME_TAG_A = "370E";
        public const string PR_ATTACH_ADDITIONAL_INFO = "370F";

        /*
         *  AB Object properties
         */

        public const string PR_DISPLAY_TYPE = "3900";
        public const string PR_TEMPLATEID = "3902";
        public const string PR_PRIMARY_CAPABILITY = "3904";


        /*
         *	Mail user properties
         */

        public const string PR_7BIT_DISPLAY_NAME = "39FF";
        public const string PR_ACCOUNT = "3A00";
        public const string PR_ACCOUNT_W = "3A00";        
        
        /// <summary>
        /// E-mail address e.g. PeterPan@neverland.com
        /// </summary>
        public const string PR_EMAIL_1 = "39FE";

        /// <summary>
        /// Second place to search for an E-mail address
        /// </summary>
        public const string PR_EMAIL_2 = "403E";
        public const string PR_ACCOUNT_A = "3A00";
        public const string PR_ALTERNATE_RECIPIENT = "3A01";
        public const string PR_CALLBACK_TELEPHONE_NUMBER = "3A02";
        public const string PR_CALLBACK_TELEPHONE_NUMBER_W = "3A02";
        public const string PR_CALLBACK_TELEPHONE_NUMBER_A = "3A02";
        public const string PR_CONVERSION_PROHIBITED = "3A03";
        public const string PR_DISCLOSE_RECIPIENTS = "3A04";
        public const string PR_GENERATION = "3A05";
        public const string PR_GENERATION_W = "3A05";
        public const string PR_GENERATION_A = "3A05";
        public const string PR_GIVEN_NAME = "3A06";
        public const string PR_GIVEN_NAME_W = "3A06";
        public const string PR_GIVEN_NAME_A = "3A06";
        public const string PR_GOVERNMENT_ID_NUMBER = "3A07";
        public const string PR_GOVERNMENT_ID_NUMBER_W = "3A07";
        public const string PR_GOVERNMENT_ID_NUMBER_A = "3A07";
        public const string PR_BUSINESS_TELEPHONE_NUMBER = "3A08";
        public const string PR_BUSINESS_TELEPHONE_NUMBER_W = "3A08";
        public const string PR_BUSINESS_TELEPHONE_NUMBER_A = "3A08";
        public const string PR_OFFICE_TELEPHONE_NUMBER = "3A08";
        public const string PR_OFFICE_TELEPHONE_NUMBER_W = "3A08";
        public const string PR_OFFICE_TELEPHONE_NUMBER_A = "3A08";
        public const string PR_HOME_TELEPHONE_NUMBER = "3A09";
        public const string PR_HOME_TELEPHONE_NUMBER_W = "3A09";
        public const string PR_HOME_TELEPHONE_NUMBER_A = "3A09";
        public const string PR_INITIALS = "3A0A";
        public const string PR_INITIALS_W = "3A0A";
        public const string PR_INITIALS_A = "3A0A";
        public const string PR_KEYWORD = "3A0B";
        public const string PR_KEYWORD_W = "3A0B";
        public const string PR_KEYWORD_A = "3A0B";
        public const string PR_LANGUAGE = "3A0C";
        public const string PR_LANGUAGE_W = "3A0C";
        public const string PR_LANGUAGE_A = "3A0C";
        public const string PR_LOCATION = "3A0D";
        public const string PR_LOCATION_W = "3A0D";
        public const string PR_LOCATION_A = "3A0D";
        public const string PR_MAIL_PERMISSION = "3A0E";
        public const string PR_MHS_COMMON_NAME = "3A0F";
        public const string PR_MHS_COMMON_NAME_W = "3A0F";
        public const string PR_MHS_COMMON_NAME_A = "3A0F";
        public const string PR_ORGANIZATIONAL_ID_NUMBER = "3A10";
        public const string PR_ORGANIZATIONAL_ID_NUMBER_W = "3A10";
        public const string PR_ORGANIZATIONAL_ID_NUMBER_A = "3A10";
        public const string PR_SURNAME = "3A11";
        public const string PR_SURNAME_W = "3A11";
        public const string PR_SURNAME_A = "3A11";
        public const string PR_ORIGINAL_ENTRYID = "3A12";
        public const string PR_ORIGINAL_DISPLAY_NAME = "3A13";
        public const string PR_ORIGINAL_DISPLAY_NAME_W = "3A13";
        public const string PR_ORIGINAL_DISPLAY_NAME_A = "3A13";
        public const string PR_ORIGINAL_SEARCH_KEY = "3A14";
        public const string PR_POSTAL_ADDRESS = "3A15";
        public const string PR_POSTAL_ADDRESS_W = "3A15";
        public const string PR_POSTAL_ADDRESS_A = "3A15";
        public const string PR_COMPANY_NAME = "3A16";
        public const string PR_COMPANY_NAME_W = "3A16";
        public const string PR_COMPANY_NAME_A = "3A16";
        public const string PR_TITLE = "3A17";
        public const string PR_TITLE_W = "3A17";
        public const string PR_TITLE_A = "3A17";
        public const string PR_DEPARTMENT_NAME = "3A18";
        public const string PR_DEPARTMENT_NAME_W = "3A18";
        public const string PR_DEPARTMENT_NAME_A = "3A18";
        public const string PR_OFFICE_LOCATION = "3A19";
        public const string PR_OFFICE_LOCATION_W = "3A19";
        public const string PR_OFFICE_LOCATION_A = "3A19";
        public const string PR_PRIMARY_TELEPHONE_NUMBER = "3A1A";
        public const string PR_PRIMARY_TELEPHONE_NUMBER_W = "3A1A";
        public const string PR_PRIMARY_TELEPHONE_NUMBER_A = "3A1A";
        public const string PR_BUSINESS2_TELEPHONE_NUMBER = "3A1B";
        public const string PR_BUSINESS2_TELEPHONE_NUMBER_W = "3A1B";
        public const string PR_BUSINESS2_TELEPHONE_NUMBER_A = "3A1B";
        public const string PR_OFFICE2_TELEPHONE_NUMBER = "3A1B";
        public const string PR_OFFICE2_TELEPHONE_NUMBER_W = "3A1B";
        public const string PR_OFFICE2_TELEPHONE_NUMBER_A = "3A1B";
        public const string PR_MOBILE_TELEPHONE_NUMBER = "3A1C";
        public const string PR_MOBILE_TELEPHONE_NUMBER_W = "3A1C";
        public const string PR_MOBILE_TELEPHONE_NUMBER_A = "3A1C";
        public const string PR_CELLULAR_TELEPHONE_NUMBER = "3A1C";
        public const string PR_CELLULAR_TELEPHONE_NUMBER_W = "3A1C";
        public const string PR_CELLULAR_TELEPHONE_NUMBER_A = "3A1C";
        public const string PR_RADIO_TELEPHONE_NUMBER = "3A1D";
        public const string PR_RADIO_TELEPHONE_NUMBER_W = "3A1D";
        public const string PR_RADIO_TELEPHONE_NUMBER_A = "3A1D";
        public const string PR_CAR_TELEPHONE_NUMBER = "3A1E";
        public const string PR_CAR_TELEPHONE_NUMBER_W = "3A1E";
        public const string PR_CAR_TELEPHONE_NUMBER_A = "3A1E";
        public const string PR_OTHER_TELEPHONE_NUMBER = "3A1F";
        public const string PR_OTHER_TELEPHONE_NUMBER_W = "3A1F";
        public const string PR_OTHER_TELEPHONE_NUMBER_A = "3A1F";
        public const string PR_TRANSMITABLE_DISPLAY_NAME = "3A20";
        public const string PR_TRANSMITABLE_DISPLAY_NAME_W = "3A20";
        public const string PR_TRANSMITABLE_DISPLAY_NAME_A = "3A20";
        public const string PR_PAGER_TELEPHONE_NUMBER = "3A21";
        public const string PR_PAGER_TELEPHONE_NUMBER_W = "3A21";
        public const string PR_PAGER_TELEPHONE_NUMBER_A = "3A21";
        public const string PR_BEEPER_TELEPHONE_NUMBER = "3A21";
        public const string PR_BEEPER_TELEPHONE_NUMBER_W = "3A21";
        public const string PR_BEEPER_TELEPHONE_NUMBER_A = "3A21";
        public const string PR_USER_CERTIFICATE = "3A22";
        public const string PR_PRIMARY_FAX_NUMBER = "3A23";
        public const string PR_PRIMARY_FAX_NUMBER_W = "3A23";
        public const string PR_PRIMARY_FAX_NUMBER_A = "3A23";
        public const string PR_BUSINESS_FAX_NUMBER = "3A24";
        public const string PR_BUSINESS_FAX_NUMBER_W = "3A24";
        public const string PR_BUSINESS_FAX_NUMBER_A = "3A24";
        public const string PR_HOME_FAX_NUMBER = "3A25";
        public const string PR_HOME_FAX_NUMBER_W = "3A25";
        public const string PR_HOME_FAX_NUMBER_A = "3A25";
        public const string PR_COUNTRY = "3A26";
        public const string PR_COUNTRY_W = "3A26";
        public const string PR_COUNTRY_A = "3A26";
        public const string PR_BUSINESS_ADDRESS_COUNTRY = "3A26";
        public const string PR_BUSINESS_ADDRESS_COUNTRY_W = "3A26";
        public const string PR_BUSINESS_ADDRESS_COUNTRY_A = "3A26";
        public const string PR_LOCALITY = "3A27";
        public const string PR_LOCALITY_W = "3A27";
        public const string PR_LOCALITY_A = "3A27";
        public const string PR_BUSINESS_ADDRESS_CITY = "3A27";
        public const string PR_BUSINESS_ADDRESS_CITY_W = "3A27";
        public const string PR_BUSINESS_ADDRESS_CITY_A = "3A27";
        public const string PR_STATE_OR_PROVINCE = "3A28";
        public const string PR_STATE_OR_PROVINCE_W = "3A28";
        public const string PR_STATE_OR_PROVINCE_A = "3A28";
        public const string PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE = "3A28";
        public const string PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE_W = "3A28";
        public const string PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE_A = "3A28";
        public const string PR_STREET_ADDRESS = "3A29";
        public const string PR_STREET_ADDRESS_W = "3A29";
        public const string PR_STREET_ADDRESS_A = "3A29";
        public const string PR_BUSINESS_ADDRESS_STREET = "3A29";
        public const string PR_BUSINESS_ADDRESS_STREET_W = "3A29";
        public const string PR_BUSINESS_ADDRESS_STREET_A = "3A29";
        public const string PR_POSTAL_CODE = "3A2A";
        public const string PR_POSTAL_CODE_W = "3A2A";
        public const string PR_POSTAL_CODE_A = "3A2A";
        public const string PR_BUSINESS_ADDRESS_POSTAL_CODE = "3A2A";
        public const string PR_BUSINESS_ADDRESS_POSTAL_CODE_W = "3A2A";
        public const string PR_BUSINESS_ADDRESS_POSTAL_CODE_A = "3A2A";
        public const string PR_POST_OFFICE_BOX = "3A2B";
        public const string PR_POST_OFFICE_BOX_W = "3A2B";
        public const string PR_POST_OFFICE_BOX_A = "3A2B";
        public const string PR_BUSINESS_ADDRESS_POST_OFFICE_BOX = "3A2B";
        public const string PR_BUSINESS_ADDRESS_POST_OFFICE_BOX_W = "3A2B";
        public const string PR_BUSINESS_ADDRESS_POST_OFFICE_BOX_A = "3A2B";
        public const string PR_TELEX_NUMBER = "3A2C";
        public const string PR_TELEX_NUMBER_W = "3A2C";
        public const string PR_TELEX_NUMBER_A = "3A2C";
        public const string PR_ISDN_NUMBER = "3A2D";
        public const string PR_ISDN_NUMBER_W = "3A2D";
        public const string PR_ISDN_NUMBER_A = "3A2D";
        public const string PR_ASSISTANT_TELEPHONE_NUMBER = "3A2E";
        public const string PR_ASSISTANT_TELEPHONE_NUMBER_W = "3A2E";
        public const string PR_ASSISTANT_TELEPHONE_NUMBER_A = "3A2E";
        public const string PR_HOME2_TELEPHONE_NUMBER = "3A2F";
        public const string PR_HOME2_TELEPHONE_NUMBER_W = "3A2F";
        public const string PR_HOME2_TELEPHONE_NUMBER_A = "3A2F";
        public const string PR_ASSISTANT = "3A30";
        public const string PR_ASSISTANT_W = "3A30";
        public const string PR_ASSISTANT_A = "3A30";
        public const string PR_SEND_RICH_INFO = "3A40";
        public const string PR_WEDDING_ANNIVERSARY = "3A41";
        public const string PR_BIRTHDAY = "3A42";
        public const string PR_HOBBIES = "3A43";
        public const string PR_HOBBIES_W = "3A43";
        public const string PR_HOBBIES_A = "3A43";
        public const string PR_MIDDLE_NAME = "3A44";
        public const string PR_MIDDLE_NAME_W = "3A44";
        public const string PR_MIDDLE_NAME_A = "3A44";
        public const string PR_DISPLAY_NAME_PREFIX = "3A45";
        public const string PR_DISPLAY_NAME_PREFIX_W = "3A45";
        public const string PR_DISPLAY_NAME_PREFIX_A = "3A45";
        public const string PR_PROFESSION = "3A46";
        public const string PR_PROFESSION_W = "3A46";
        public const string PR_PROFESSION_A = "3A46";

        public const string PR_PREFERRED_BY_NAME = "3A47";
        public const string PR_PREFERRED_BY_NAME_W = "3A47";
        public const string PR_PREFERRED_BY_NAME_A = "3A47";

        public const string PR_SPOUSE_NAME = "3A48";
        public const string PR_SPOUSE_NAME_W = "3A48";
        public const string PR_SPOUSE_NAME_A = "3A48";

        public const string PR_COMPUTER_NETWORK_NAME = "3A49";
        public const string PR_COMPUTER_NETWORK_NAME_W = "3A49";
        public const string PR_COMPUTER_NETWORK_NAME_A = "3A49";

        public const string PR_CUSTOMER_ID = "3A4A";
        public const string PR_CUSTOMER_ID_W = "3A4A";
        public const string PR_CUSTOMER_ID_A = "3A4A";

        public const string PR_TTYTDD_PHONE_NUMBER = "3A4B";
        public const string PR_TTYTDD_PHONE_NUMBER_W = "3A4B";
        public const string PR_TTYTDD_PHONE_NUMBER_A = "3A4B";

        public const string PR_FTP_SITE = "3A4C";
        public const string PR_FTP_SITE_W = "3A4C";
        public const string PR_FTP_SITE_A = "3A4C";

        public const string PR_GENDER = "3A4D";

        public const string PR_MANAGER_NAME = "3A4E";
        public const string PR_MANAGER_NAME_W = "3A4E";
        public const string PR_MANAGER_NAME_A = "3A4E";

        public const string PR_NICKNAME = "3A4F";
        public const string PR_NICKNAME_W = "3A4F";
        public const string PR_NICKNAME_A = "3A4F";

        public const string PR_PERSONAL_HOME_PAGE = "3A50";
        public const string PR_PERSONAL_HOME_PAGE_W = "3A50";
        public const string PR_PERSONAL_HOME_PAGE_A = "3A50";

        public const string PR_BUSINESS_HOME_PAGE = "3A51";
        public const string PR_BUSINESS_HOME_PAGE_W = "3A51";
        public const string PR_BUSINESS_HOME_PAGE_A = "3A51";

        public const string PR_CONTACT_VERSION = "3A52";
        public const string PR_CONTACT_ENTRYIDS = "3A53";

        public const string PR_CONTACT_ADDRTYPES = "3A54";
        public const string PR_CONTACT_ADDRTYPES_W = "3A54";
        public const string PR_CONTACT_ADDRTYPES_A = "3A54";

        public const string PR_CONTACT_DEFAULT_ADDRESS_INDEX = "3A55";

        public const string PR_CONTACT_EMAIL_ADDRESSES = "3A56";
        public const string PR_CONTACT_EMAIL_ADDRESSES_W = "3A56";
        public const string PR_CONTACT_EMAIL_ADDRESSES_A = "3A56";

        public const string PR_COMPANY_MAIN_PHONE_NUMBER = "3A57";
        public const string PR_COMPANY_MAIN_PHONE_NUMBER_W = "3A57";
        public const string PR_COMPANY_MAIN_PHONE_NUMBER_A = "3A57";

        public const string PR_CHILDRENS_NAMES = "3A58";
        public const string PR_CHILDRENS_NAMES_W = "3A58";
        public const string PR_CHILDRENS_NAMES_A = "3A58";

        public const string PR_HOME_ADDRESS_CITY = "3A59";
        public const string PR_HOME_ADDRESS_CITY_W = "3A59";
        public const string PR_HOME_ADDRESS_CITY_A = "3A59";

        public const string PR_HOME_ADDRESS_COUNTRY = "3A5A";
        public const string PR_HOME_ADDRESS_COUNTRY_W = "3A5A";
        public const string PR_HOME_ADDRESS_COUNTRY_A = "3A5A";

        public const string PR_HOME_ADDRESS_POSTAL_CODE = "3A5B";
        public const string PR_HOME_ADDRESS_POSTAL_CODE_W = "3A5B";
        public const string PR_HOME_ADDRESS_POSTAL_CODE_A = "3A5B";

        public const string PR_HOME_ADDRESS_STATE_OR_PROVINCE = "3A5C";
        public const string PR_HOME_ADDRESS_STATE_OR_PROVINCE_W = "3A5C";
        public const string PR_HOME_ADDRESS_STATE_OR_PROVINCE_A = "3A5C";

        public const string PR_HOME_ADDRESS_STREET = "3A5D";
        public const string PR_HOME_ADDRESS_STREET_W = "3A5D";
        public const string PR_HOME_ADDRESS_STREET_A = "3A5D";

        public const string PR_HOME_ADDRESS_POST_OFFICE_BOX = "3A5E";
        public const string PR_HOME_ADDRESS_POST_OFFICE_BOX_W = "3A5E";
        public const string PR_HOME_ADDRESS_POST_OFFICE_BOX_A = "3A5E";

        public const string PR_OTHER_ADDRESS_CITY = "3A5F";
        public const string PR_OTHER_ADDRESS_CITY_W = "3A5F";
        public const string PR_OTHER_ADDRESS_CITY_A = "3A5F";

        public const string PR_OTHER_ADDRESS_COUNTRY = "3A60";
        public const string PR_OTHER_ADDRESS_COUNTRY_W = "3A60";
        public const string PR_OTHER_ADDRESS_COUNTRY_A = "3A60";

        public const string PR_OTHER_ADDRESS_POSTAL_CODE = "3A61";
        public const string PR_OTHER_ADDRESS_POSTAL_CODE_W = "3A61";
        public const string PR_OTHER_ADDRESS_POSTAL_CODE_A = "3A61";

        public const string PR_OTHER_ADDRESS_STATE_OR_PROVINCE = "3A62";
        public const string PR_OTHER_ADDRESS_STATE_OR_PROVINCE_W = "3A62";
        public const string PR_OTHER_ADDRESS_STATE_OR_PROVINCE_A = "3A62";

        public const string PR_OTHER_ADDRESS_STREET = "3A63";
        public const string PR_OTHER_ADDRESS_STREET_W = "3A63";
        public const string PR_OTHER_ADDRESS_STREET_A = "3A63";

        public const string PR_OTHER_ADDRESS_POST_OFFICE_BOX = "3A64";
        public const string PR_OTHER_ADDRESS_POST_OFFICE_BOX_W = "3A64";
        public const string PR_OTHER_ADDRESS_POST_OFFICE_BOX_A = "3A64";


        /*
         *	Profile section properties
         */

        public const string PR_STORE_PROVIDERS = "3D00";
        public const string PR_AB_PROVIDERS = "3D01";
        public const string PR_TRANSPORT_PROVIDERS = "3D02";

        public const string PR_DEFAULT_PROFILE = "3D04";
        public const string PR_AB_SEARCH_PATH = "3D05";
        public const string PR_AB_DEFAULT_DIR = "3D06";
        public const string PR_AB_DEFAULT_PAB = "3D07";

        public const string PR_FILTERING_HOOKS = "3D08";
        public const string PR_SERVICE_NAME = "3D09";
        public const string PR_SERVICE_NAME_W = "3D09";
        public const string PR_SERVICE_NAME_A = "3D09";
        public const string PR_SERVICE_DLL_NAME = "3D0A";
        public const string PR_SERVICE_DLL_NAME_W = "3D0A";
        public const string PR_SERVICE_DLL_NAME_A = "3D0A";
        public const string PR_SERVICE_ENTRY_NAME = "3D0B";
        public const string PR_SERVICE_UID = "3D0C";
        public const string PR_SERVICE_EXTRA_UIDS = "3D0D";
        public const string PR_SERVICES = "3D0E";
        public const string PR_SERVICE_SUPPORT_FILES = "3D0F";
        public const string PR_SERVICE_SUPPORT_FILES_W = "3D0F";
        public const string PR_SERVICE_SUPPORT_FILES_A = "3D0F";
        public const string PR_SERVICE_DELETE_FILES = "3D10";
        public const string PR_SERVICE_DELETE_FILES_W = "3D10";
        public const string PR_SERVICE_DELETE_FILES_A = "3D10";
        public const string PR_AB_SEARCH_PATH_UPDATE = "3D11";
        public const string PR_PROFILE_NAME = "3D12";
        public const string PR_PROFILE_NAME_A = "3D12";
        public const string PR_PROFILE_NAME_W = "3D12";

        /*
         *	Status object properties
         */

        public const string PR_IDENTITY_DISPLAY = "3E00";
        public const string PR_IDENTITY_DISPLAY_W = "3E00";
        public const string PR_IDENTITY_DISPLAY_A = "3E00";
        public const string PR_IDENTITY_ENTRYID = "3E01";
        public const string PR_RESOURCE_METHODS = "3E02";
        public const string PR_RESOURCE_TYPE = "3E03";
        public const string PR_STATUS_CODE = "3E04";
        public const string PR_IDENTITY_SEARCH_KEY = "3E05";
        public const string PR_OWN_STORE_ENTRYID = "3E06";
        public const string PR_RESOURCE_PATH = "3E07";
        public const string PR_RESOURCE_PATH_W = "3E07";
        public const string PR_RESOURCE_PATH_A = "3E07";
        public const string PR_STATUS_STRING = "3E08";
        public const string PR_STATUS_STRING_W = "3E08";
        public const string PR_STATUS_STRING_A = "3E08";
        public const string PR_X400_DEFERRED_DELIVERY_CANCEL = "3E09";
        public const string PR_HEADER_FOLDER_ENTRYID = "3E0A";
        public const string PR_REMOTE_PROGRESS = "3E0B";
        public const string PR_REMOTE_PROGRESS_TEXT = "3E0C";
        public const string PR_REMOTE_PROGRESS_TEXT_W = "3E0C";
        public const string PR_REMOTE_PROGRESS_TEXT_A = "3E0C";
        public const string PR_REMOTE_VALIDATE_OK = "3E0D";

        /*
         * Display table properties
         */

        public const string PR_CONTROL_FLAGS = "3F00";
        public const string PR_CONTROL_STRUCTURE = "3F01";
        public const string PR_CONTROL_TYPE = "3F02";
        public const string PR_DELTAX = "3F03";
        public const string PR_DELTAY = "3F04";
        public const string PR_XPOS = "3F05";
        public const string PR_YPOS = "3F06";
        public const string PR_CONTROL_ID = "3F07";
        public const string PR_INITIAL_DETAILS_PANE = "3F08";

        /*
         * Secure property id range
         */

        public const string PROP_ID_SECURE_MIN = "67F0";
        public const string PROP_ID_SECURE_MAX = "67FF";

        /* MAPITAGS_H */

        /// <summary>
        ///     Specifies the color to be used when displaying a Calendar object
        /// </summary>
        public const string PidNameKeywords = "PidNameKeywords";
        #endregion

        #region Mapi tag types
        /// <summary>
        ///     (Reserved for interface use) type doesn't matter to caller
        /// </summary>
        public const ushort PT_UNSPECIFIED = 0;

        /// <summary>
        ///     NULL property value
        /// </summary>
        public const ushort PT_NULL = 1;

        /// <summary>
        ///     Signed 16-bit value
        /// </summary>
        public const ushort PT_I2 = 2;

        /// <summary>
        ///     Signed 32-bit value
        /// </summary>
        public const ushort PT_LONG = 3;

        /// <summary>
        ///     4-byte floating point
        /// </summary>
        public const ushort PT_R4 = 4;

        /// <summary>
        ///     Floating point double
        /// </summary>
        public const ushort PT_DOUBLE = 5;

        /// <summary>
        ///     Signed 64-bit int (decimal w/4 digits right of decimal pt)
        /// </summary>
        public const ushort PT_CURRENCY = 6;

        /// <summary>
        ///     Application time
        /// </summary>
        public const ushort PT_APPTIME = 7;

        /// <summary>
        ///     32-bit error value
        /// </summary>
        public const ushort PT_ERROR = 10;

        /// <summary>
        ///     16-bit boolean (non-zero true)
        /// </summary>
        public const ushort PT_BOOLEAN = 11;

        /// <summary>
        ///     Embedded object in a property
        /// </summary>
        public const ushort PT_OBJECT = 13;

        /// <summary>
        ///     8-byte signed integer
        /// </summary>
        public const ushort PT_I8 = 20;

        /// <summary>
        ///     Null terminated 8-bit character string
        /// </summary>
        public const ushort PT_STRING8 = 30;

        /// <summary>
        ///     Null terminated Unicode string
        /// </summary>
        public const ushort PT_UNICODE = 31;

        /// <summary>
        ///     FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601
        /// </summary>
        public const ushort PT_SYSTIME = 64;

        /// <summary>
        ///     OLE GUID
        /// </summary>
        public const ushort PT_CLSID = 72;

        /// <summary>
        ///     Uninterpreted (counted byte array)
        /// </summary>
        public const ushort PT_BINARY = 258;

        /// <summary>
        ///     Multi-view unicode string
        /// </summary>
        public const ushort PT_MV_UNICODE = 4127;
        #endregion

        #region Stream constants
        /// <summary>
        ///     Storage prefix tag
        /// </summary>
        public const string RecipStoragePrefix = "__recip_version1.0_#";

        /// <summary>
        ///     Prefix that is placed before an attachment tag
        /// </summary>
        public const string AttachStoragePrefix = "__attach_version1.0_#";

        /// <summary>
        ///     Sub storage version 1.0 streams
        /// </summary>
        public const string SubStgVersion1 = "__substg1.0";

        /// <summary>
        ///     Stream that contains the internet E-mail headers
        /// </summary>
        public const string HeaderStreamName = "__substg1.0_007D001F";

        /// <summary>
        ///     The stream that contains all the MAPI properties
        /// </summary>
        public const string PropertiesStream = "__properties_version1.0";

        /// <summary>
        ///     Contains the streams needed to perform named property mapping
        /// </summary>
        public const string NameIdStorage = "__nameid_version1.0";

        /// <summary>
        ///     The stream with the name properties are always in stream "__substg1.0_00030102"
        /// </summary>
        public const string NameIdStorageMappingStream = "__substg1.0_00030102";
        public const string NameIdStorageMappingStream2 = "__substg1.0_00040102";

        /// <summary>
        ///     Stream properties begin for header or top
        /// </summary>
        public const int PropertiesStreamHeaderTop = 32;

        // Stream properties begin for embeded
        public const int PropertiesStreamHeaderEmbeded = 24;

        /// <summary>
        ///     Stream properties begin for attachments or recipients
        /// </summary>
        public const int PropertiesStreamHeaderAttachOrRecip = 8;
        #endregion
        
        #region Attachment type constants
        /// <summary>
        ///     There is no attachment
        /// </summary>
        public const int NO_ATTACHMENT = 0;

        public const int ATTACH_BY_VALUE = 1;
        public const int ATTACH_BY_REFERENCE = 2;
        public const int ATTACH_BY_REF_RESOLVE = 3;
        public const int ATTACH_BY_REF_ONLY = 4;

        /// <summary>
        ///     The attachment is a msg file
        /// </summary>
        public const int ATTACH_EMBEDDED_MSG = 5;

        /// <summary>
        ///     The attachment in an OLE object
        /// </summary>
        public const int ATTACH_OLE = 6;
        #endregion

        #region MAPI TO, CC and BCC contstants
        /// <summary>
        ///     E-mail To address
        /// </summary>
        public const int MAPI_TO = 1;

        /// <summary>
        ///     E-mail From address
        /// </summary>
        public const int MAPI_CC = 2;

        /// <summary>
        ///     E-mail BCC address
        /// </summary>
        public const int MAPI_BCC = 3;
        #endregion

        #region Flag constants
        /// <summary>
        ///     E-mail follow up flag (named property)
        /// </summary>
        public const string FlagRequest = "8530";

        /// <summary>
        ///     Specifies the flag state of the message object; Not present, 1 = Completed, 2 = Flagged.
        ///     Only available from Outlook 2007 and up.
        /// </summary>
        public const string PR_FLAG_STATUS = "1090";

        /// <summary>
        ///     Contains the date when the task was completed. Only filled when <see cref="TaskComplete" /> is true.
        ///     Only available from Outlook 2007 and up.
        /// </summary>
        public const string PR_FLAG_COMPLETE_TIME = "1091";
        #endregion

        #region Task constants
        /// <summary>
        ///     <see cref="TaskStatus" /> of the task (named property)
        /// </summary>
        public const string TaskStatus = "8006";

        /// <summary>
        ///     Start date of the task (named property)
        /// </summary>
        public const string TaskStartDate = "8104";

        /// <summary>
        ///     Due date of the task (named property)
        /// </summary>
        public const string TaskDueDate = "8105";

        /// <summary>
        ///     True when the task is complete (named property)
        /// </summary>
        public const string TaskComplete = "811C";
        #endregion

        #region Appointment constants
        /// <summary>
        ///     Appointment location (named property)
        /// </summary>
        public const string Location = "8208";

        /// <summary>
        ///     Appointment reccurence type (named property)
        /// </summary>
        public const string ReccurrenceType = "8231";

        /// <summary>
        ///     Appointment reccurence pattern (named property)
        /// </summary>
        public const string ReccurrencePattern = "8232";

        /// <summary>
        ///     Appointment start time (greenwich time) (named property)
        /// </summary>
        public const string AppointmentStartWhole = "820D";

        /// <summary>
        ///     Appointment end time (greenwich time) (named property)
        /// </summary>
        public const string AppointmentEndWhole = "820E";
        #endregion

        /// <summary>
        /// E-mail address of the sender e.g. PeterPan@neverland.com (named property)
        /// </summary>
        public const string PR_SENDER_EMAIL_ADDRESS_2 = "8012";
        // ReSharper restore InconsistentNaming
    }
}