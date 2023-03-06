//
// PropertyTag.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2022 .NET Foundation and Contributors
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

using MsgReader.Tnef.Enums;

namespace MsgReader.Tnef
{
    /// <summary>
    /// A TNEF property tag.
    /// </summary>
    /// <remarks>
    /// A TNEF property tag.
    /// </remarks>
    public struct PropertyTag
    {
        /// <summary>
        /// The MAPI property PR_AB_DEFAULT_DIR.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AB_DEFAULT_DIR.
        /// </remarks>
        public static readonly PropertyTag AbDefaultDir = new PropertyTag(PropertyId.AbDefaultDir, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_AB_DEFAULT_PAB.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AB_DEFAULT_PAD.
        /// </remarks>
        public static readonly PropertyTag AbDefaultPab = new PropertyTag(PropertyId.AbDefaultPab, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_AB_PROVIDER_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AB_PROVIDER_ID.
        /// </remarks>
        public static readonly PropertyTag AbProviderId = new PropertyTag(PropertyId.AbProviderId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_AB_PROVIDERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AB_PROVIDERS.
        /// </remarks>
        public static readonly PropertyTag AbProviders = new PropertyTag(PropertyId.AbProviders, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_AB_SEARCH_PATH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AB_SEARCH_PATH.
        /// </remarks>
        public static readonly PropertyTag AbSearchPath = new PropertyTag(PropertyId.AbSearchPath, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_AB_SEARCH_PATH_UPDATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AB_SEARCH_PATH_UPDATE.
        /// </remarks>
        public static readonly PropertyTag AbSearchPathUpdate = new PropertyTag(PropertyId.AbSearchPathUpdate, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ACCESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ACCESS.
        /// </remarks>
        public static readonly PropertyTag Access = new PropertyTag(PropertyId.Access, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ACCESS_LEVEL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ACCESS_LEVEL.
        /// </remarks>
        public static readonly PropertyTag AccessLevel = new PropertyTag(PropertyId.AccessLevel, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ACCOUNT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ACCOUNT.
        /// </remarks>
        public static readonly PropertyTag AccountA = new PropertyTag(PropertyId.Account, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ACCOUNT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ACCOUNT.
        /// </remarks>
        public static readonly PropertyTag AccountW = new PropertyTag(PropertyId.Account, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ACKNOWLEDGEMENT_MODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ACKNOWLEDGEMENT_MODE.
        /// </remarks>
        public static readonly PropertyTag AcknowledgementMode = new PropertyTag(PropertyId.AcknowledgementMode, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag AddrtypeA = new PropertyTag(PropertyId.Addrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag AddrtypeW = new PropertyTag(PropertyId.Addrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ALTERNATE_RECIPIENT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ALTERNATE_RECIPIENT.
        /// </remarks>
        public static readonly PropertyTag AlternateRecipient = new PropertyTag(PropertyId.AlternateRecipient, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ALTERNATE_RECIPIENT_ALLOWED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ALTERNATE_RECIPIENT_ALLOWED.
        /// </remarks>
        public static readonly PropertyTag AlternateRecipientAllowed = new PropertyTag(PropertyId.AlternateRecipientAllowed, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_ANR.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ANR.
        /// </remarks>
        public static readonly PropertyTag AnrA = new PropertyTag(PropertyId.Anr, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ANR.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ANR.
        /// </remarks>
        public static readonly PropertyTag AnrW = new PropertyTag(PropertyId.Anr, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ASSISTANT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ASSISTANT.
        /// </remarks>
        public static readonly PropertyTag AssistantA = new PropertyTag(PropertyId.Assistant, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ASSISTANT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ASSISTANT.
        /// </remarks>
        public static readonly PropertyTag AssistantW = new PropertyTag(PropertyId.Assistant, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag AssistantTelephoneNumberA = new PropertyTag(PropertyId.AssistantTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag AssistantTelephoneNumberW = new PropertyTag(PropertyId.AssistantTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ASSOC_CONTENT_COUNT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ASSOC_CONTENT_COUNT.
        /// </remarks>
        public static readonly PropertyTag AssocContentCount = new PropertyTag(PropertyId.AssocContentCount, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ATTACH_ADDITIONAL_INFO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_ADDITIONAL_INFO.
        /// </remarks>
        public static readonly PropertyTag AttachAdditionalInfo = new PropertyTag(PropertyId.AttachAdditionalInfo, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ATTACH_CONTENT_BASE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_CONTENT_BASE.
        /// </remarks>
        public static readonly PropertyTag AttachContentBaseA = new PropertyTag(PropertyId.AttachContentBase, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_CONTENT_BASE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_CONTENT_BASE.
        /// </remarks>
        public static readonly PropertyTag AttachContentBaseW = new PropertyTag(PropertyId.AttachContentBase, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_CONTENT_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_CONTENT_ID.
        /// </remarks>
        public static readonly PropertyTag AttachContentIdA = new PropertyTag(PropertyId.AttachContentId, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_CONTENT_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_CONTENT_ID.
        /// </remarks>
        public static readonly PropertyTag AttachContentIdW = new PropertyTag(PropertyId.AttachContentId, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_CONTENT_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_CONTENT_LOCATION.
        /// </remarks>
        public static readonly PropertyTag AttachContentLocationA = new PropertyTag(PropertyId.AttachContentLocation, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_CONTENT_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_CONTENT_LOCATION.
        /// </remarks>
        public static readonly PropertyTag AttachContentLocationW = new PropertyTag(PropertyId.AttachContentLocation, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_DATA.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_DATA.
        /// </remarks>
        public static readonly PropertyTag AttachDataBin = new PropertyTag(PropertyId.AttachData, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ATTACH_DATA.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_DATA.
        /// </remarks>
        public static readonly PropertyTag AttachDataObj = new PropertyTag(PropertyId.AttachData, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_ATTACH_DISPOSITION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_DISPOSITION.
        /// </remarks>
        public static readonly PropertyTag AttachDispositionA = new PropertyTag(PropertyId.AttachDisposition, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_DISPOSITION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_DISPOSITION.
        /// </remarks>
        public static readonly PropertyTag AttachDispositionW = new PropertyTag(PropertyId.AttachDisposition, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_ENCODING.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_ENCODING.
        /// </remarks>
        public static readonly PropertyTag AttachEncoding = new PropertyTag(PropertyId.AttachEncoding, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ATTACH_EXTENSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_EXTENSION.
        /// </remarks>
        public static readonly PropertyTag AttachExtensionA = new PropertyTag(PropertyId.AttachExtension, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_EXTENSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_EXTENSION.
        /// </remarks>
        public static readonly PropertyTag AttachExtensionW = new PropertyTag(PropertyId.AttachExtension, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_FILENAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_FILENAME.
        /// </remarks>
        public static readonly PropertyTag AttachFilenameA = new PropertyTag(PropertyId.AttachFilename, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_FILENAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_FILENAME.
        /// </remarks>
        public static readonly PropertyTag AttachFilenameW = new PropertyTag(PropertyId.AttachFilename, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_FLAGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_FLAGS.
        /// </remarks>
        public static readonly PropertyTag AttachFlags = new PropertyTag(PropertyId.AttachFlags, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_ATTACH_LONG_FILENAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_LONG_FILENAME.
        /// </remarks>
        public static readonly PropertyTag AttachLongFilenameA = new PropertyTag(PropertyId.AttachLongFilename, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_LONG_FILENAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_LONG_FILENAME.
        /// </remarks>
        public static readonly PropertyTag AttachLongFilenameW = new PropertyTag(PropertyId.AttachLongFilename, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_LONG_PATHNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_LONG_PATHNAME.
        /// </remarks>
        public static readonly PropertyTag AttachLongPathnameA = new PropertyTag(PropertyId.AttachLongPathname, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_LONG_PATHNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_LONG_PATHNAME.
        /// </remarks>
        public static readonly PropertyTag AttachLongPathnameW = new PropertyTag(PropertyId.AttachLongPathname, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACHMENT_CONTACTPHOTO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACHMENT_CONTACTPHOTO.
        /// </remarks>
        public static readonly PropertyTag AttachmentContactPhoto = new PropertyTag(PropertyId.AttachmentContactPhoto, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_ATTACHMENT_FLAGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACHMENT_FLAGS.
        /// </remarks>
        public static readonly PropertyTag AttachmentFlags = new PropertyTag(PropertyId.AttachmentFlags, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ATTACHMENT_HIDDEN.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACHMENT_HIDDEN.
        /// </remarks>
        public static readonly PropertyTag AttachmentHidden = new PropertyTag(PropertyId.AttachmentHidden, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_ATTACHMENT_LINKID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACHMENT_LINKID.
        /// </remarks>
        public static readonly PropertyTag AttachmentLinkId = new PropertyTag(PropertyId.AttachmentLinkId, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ATTACHMENT_X400_PARAMETERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACHMENT_X400_PARAMETERS.
        /// </remarks>
        public static readonly PropertyTag AttachmentX400Parameters = new PropertyTag(PropertyId.AttachmentX400Parameters, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ATTACH_METHOD.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_METHOD.
        /// </remarks>
        public static readonly PropertyTag AttachMethod = new PropertyTag(PropertyId.AttachMethod, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ATTACH_MIME_SEQUENCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_MIME_SEQUENCE.
        /// </remarks>
        public static readonly PropertyTag AttachMimeSequence = new PropertyTag(PropertyId.AttachMimeSequence, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_ATTACH_MIME_TAG.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_MIME_TAG.
        /// </remarks>
        public static readonly PropertyTag AttachMimeTagA = new PropertyTag(PropertyId.AttachMimeTag, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_MIME_TAG.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_MIME_TAG.
        /// </remarks>
        public static readonly PropertyTag AttachMimeTagW = new PropertyTag(PropertyId.AttachMimeTag, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_NETSCAPE_MAC_INFO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_NETSCAPE_MAC_INFO.
        /// </remarks>
        public static readonly PropertyTag AttachNetscapeMacInfo = new PropertyTag(PropertyId.AttachNetscapeMacInfo, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_ATTACH_NUM.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_NUM.
        /// </remarks>
        public static readonly PropertyTag AttachNum = new PropertyTag(PropertyId.AttachNum, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ATTACH_PATHNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_PATHNAME.
        /// </remarks>
        public static readonly PropertyTag AttachPathnameA = new PropertyTag(PropertyId.AttachPathname, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_PATHNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_PATHNAME.
        /// </remarks>
        public static readonly PropertyTag AttachPathnameW = new PropertyTag(PropertyId.AttachPathname, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ATTACH_RENDERING.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_RENDERING.
        /// </remarks>
        public static readonly PropertyTag AttachRendering = new PropertyTag(PropertyId.AttachRendering, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ATTACH_SIZE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_SIZE.
        /// </remarks>
        public static readonly PropertyTag AttachSize = new PropertyTag(PropertyId.AttachSize, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ATTACH_TAG.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_TAG.
        /// </remarks>
        public static readonly PropertyTag AttachTag = new PropertyTag(PropertyId.AttachTag, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ATTACH_TRANSPORT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_TRANSPORT_NAME.
        /// </remarks>
        public static readonly PropertyTag AttachTransportNameA = new PropertyTag(PropertyId.AttachTransportName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ATTACH_TRANSPORT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ATTACH_TRANSPORT_NAME.
        /// </remarks>
        public static readonly PropertyTag AttachTransportNameW = new PropertyTag(PropertyId.AttachTransportName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_AUTHORIZING_USERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AUTHORIZING_USERS.
        /// </remarks>
        public static readonly PropertyTag AuthorizingUsers = new PropertyTag(PropertyId.AuthorizingUsers, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_AUTOFORWARDED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AUTOFORWARDED.
        /// </remarks>
        public static readonly PropertyTag AutoForwarded = new PropertyTag(PropertyId.AutoForwarded, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_AUTOFORWARDING_COMMENT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AUTOFORWARDING_COMMENT.
        /// </remarks>
        public static readonly PropertyTag AutoForwardingCommentA = new PropertyTag(PropertyId.AutoForwardingComment, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_AUTOFORWARDING_COMMENT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AUTOFORWARDING_COMMENT.
        /// </remarks>
        public static readonly PropertyTag AutoForwardingCommentW = new PropertyTag(PropertyId.AutoForwardingComment, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_AUTORESPONSE_SUPPRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_AUTORESPONSE_SUPPRESS.
        /// </remarks>
        public static readonly PropertyTag AutoResponseSuppress = new PropertyTag(PropertyId.AutoResponseSuppress, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag BeeperTelephoneNumberA = new PropertyTag(PropertyId.BeeperTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag BeeperTelephoneNumberW = new PropertyTag(PropertyId.BeeperTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BIRTHDAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BIRTHDAY.
        /// </remarks>
        public static readonly PropertyTag Birthday = new PropertyTag(PropertyId.Birthday, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_BODY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY.
        /// </remarks>
        public static readonly PropertyTag BodyA = new PropertyTag(PropertyId.Body, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BODY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY.
        /// </remarks>
        public static readonly PropertyTag BodyW = new PropertyTag(PropertyId.Body, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BODY_CONTENT_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_CONTENT_ID.
        /// </remarks>
        public static readonly PropertyTag BodyContentIdA = new PropertyTag(PropertyId.BodyContentId, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BODY_CONTENT_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_CONTENT_ID.
        /// </remarks>
        public static readonly PropertyTag BodyContentIdW = new PropertyTag(PropertyId.BodyContentId, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BODY_CONTENT_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_CONTENT_LOCATION.
        /// </remarks>
        public static readonly PropertyTag BodyContentLocationA = new PropertyTag(PropertyId.BodyContentLocation, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BODY_CONTENT_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_CONTENT_LOCATION.
        /// </remarks>
        public static readonly PropertyTag BodyContentLocationW = new PropertyTag(PropertyId.BodyContentLocation, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BODY_CRC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_CRC.
        /// </remarks>
        public static readonly PropertyTag BodyCrc = new PropertyTag(PropertyId.BodyCrc, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_BODY_HTML.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_HTML.
        /// </remarks>
        public static readonly PropertyTag BodyHtmlA = new PropertyTag(PropertyId.BodyHtml, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BODY_HTML.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_HTML.
        /// </remarks>
        public static readonly PropertyTag BodyHtmlB = new PropertyTag(PropertyId.BodyHtml, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_BODY_HTML.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BODY_HTML.
        /// </remarks>
        public static readonly PropertyTag BodyHtmlW = new PropertyTag(PropertyId.BodyHtml, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Business2TelephoneNumberA = new PropertyTag(PropertyId.Business2TelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Business2TelephoneNumberAMv = new PropertyTag(PropertyId.Business2TelephoneNumber, PropertyType.String8, true);

        /// <summary>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Business2TelephoneNumberW = new PropertyTag(PropertyId.Business2TelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Business2TelephoneNumberWMv = new PropertyTag(PropertyId.Business2TelephoneNumber, PropertyType.Unicode, true);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_CITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_CITY.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressCityA = new PropertyTag(PropertyId.BusinessAddressCity, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_CITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_CITY.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressCityW = new PropertyTag(PropertyId.BusinessAddressCity, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressCountryA = new PropertyTag(PropertyId.BusinessAddressCountry, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressCountryW = new PropertyTag(PropertyId.BusinessAddressCountry, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressPostalCodeA = new PropertyTag(PropertyId.BusinessAddressPostalCode, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressPostalCodeW = new PropertyTag(PropertyId.BusinessAddressPostalCode, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressStreetA = new PropertyTag(PropertyId.BusinessAddressStreet, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
        /// </remarks>
        public static readonly PropertyTag BusinessAddressStreetW = new PropertyTag(PropertyId.BusinessAddressStreet, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BUSINESS_FAX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_FAX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag BusinessFaxNumberA = new PropertyTag(PropertyId.BusinessFaxNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BUSINESS_FAX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_FAX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag BusinessFaxNumberW = new PropertyTag(PropertyId.BusinessFaxNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_BUSINESS_HOME_PAGE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_HOME_PAGE.
        /// </remarks>
        public static readonly PropertyTag BusinessHomePageA = new PropertyTag(PropertyId.BusinessHomePage, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_BUSINESS_HOME_PAGE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_BUSINESS_HOME_PAGE.
        /// </remarks>
        public static readonly PropertyTag BusinessHomePageW = new PropertyTag(PropertyId.BusinessHomePage, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag CallbackTelephoneNumberA = new PropertyTag(PropertyId.CallbackTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag CallbackTelephoneNumberW = new PropertyTag(PropertyId.CallbackTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CAR_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CAR_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag CarTelephoneNumberA = new PropertyTag(PropertyId.CarTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_CAR_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CAR_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag CarTelephoneNumberW = new PropertyTag(PropertyId.CarTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CHILDRENS_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CHILDRENS_NAMES.
        /// </remarks>
        public static readonly PropertyTag ChildrensNamesA = new PropertyTag(PropertyId.ChildrensNames, PropertyType.String8, true);

        /// <summary>
        /// The MAPI property PR_CHILDRENS_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CHILDRENS_NAMES.
        /// </remarks>
        public static readonly PropertyTag ChildrensNamesW = new PropertyTag(PropertyId.ChildrensNames, PropertyType.Unicode, true);

        /// <summary>
        /// The MAPI property PR_CLIENT_SUBMIT_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CLIENT_SUBMIT_TIME.
        /// </remarks>
        public static readonly PropertyTag ClientSubmitTime = new PropertyTag(PropertyId.ClientSubmitTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_COMMENT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMMENT.
        /// </remarks>
        public static readonly PropertyTag CommentA = new PropertyTag(PropertyId.Comment, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_COMMENT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMMENT.
        /// </remarks>
        public static readonly PropertyTag CommentW = new PropertyTag(PropertyId.Comment, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_COMMON_VIEWS_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMMON_VIEWS_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag CommonViewsEntryId = new PropertyTag(PropertyId.CommonViewsEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag CompanyMainPhoneNumberA = new PropertyTag(PropertyId.CompanyMainPhoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag CompanyMainPhoneNumberW = new PropertyTag(PropertyId.CompanyMainPhoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_COMPANY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMPANY_NAME.
        /// </remarks>
        public static readonly PropertyTag CompanyNameA = new PropertyTag(PropertyId.CompanyName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_COMPANY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMPANY_NAME.
        /// </remarks>
        public static readonly PropertyTag CompanyNameW = new PropertyTag(PropertyId.CompanyName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_COMPUTER_NETWORK_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMPUTER_NETWORK_NAME.
        /// </remarks>
        public static readonly PropertyTag ComputerNetworkNameA = new PropertyTag(PropertyId.ComputerNetworkName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_COMPUTER_NETWORK_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COMPUTER_NETWORK_NAME.
        /// </remarks>
        public static readonly PropertyTag ComputerNetworkNameW = new PropertyTag(PropertyId.ComputerNetworkName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CONTACT_ADDRTYPES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTACT_ADDRTYPES.
        /// </remarks>
        public static readonly PropertyTag ContactAddrtypesA = new PropertyTag(PropertyId.ContactAddrtypes, PropertyType.String8, true);

        /// <summary>
        /// The MAPI property PR_CONTACT_ADDRTYPES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTACT_ADDRTYPES.
        /// </remarks>
        public static readonly PropertyTag ContactAddrtypesW = new PropertyTag(PropertyId.ContactAddrtypes, PropertyType.Unicode, true);

        /// <summary>
        /// The MAPI property PR_CONTACT_DEFAULT_ADDRESS_INDEX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTACT_DEFAULT_ADDRESS_INDEX.
        /// </remarks>
        public static readonly PropertyTag ContactDefaultAddressIndex = new PropertyTag(PropertyId.ContactDefaultAddressIndex, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
        /// </remarks>
        public static readonly PropertyTag ContactEmailAddressesA = new PropertyTag(PropertyId.ContactEmailAddresses, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
        /// </remarks>
        public static readonly PropertyTag ContactEmailAddressesW = new PropertyTag(PropertyId.ContactEmailAddresses, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CONTACT_ENTRYIDS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTACT_ENTRYIDS.
        /// </remarks>
        public static readonly PropertyTag ContactEntryIds = new PropertyTag(PropertyId.ContactEntryIds, PropertyType.Binary, true);

        /// <summary>
        /// The MAPI property PR_CONTACT_VERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTACT_VERSION.
        /// </remarks>
        public static readonly PropertyTag ContactVersion = new PropertyTag(PropertyId.ContactVersion, PropertyType.ClassId);

        /// <summary>
        /// The MAPI property PR_CONTAINER_CLASS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTAINER_CLASS.
        /// </remarks>
        public static readonly PropertyTag ContainerClassA = new PropertyTag(PropertyId.ContainerClass, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_CONTAINER_CLASS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTAINER_CLASS.
        /// </remarks>
        public static readonly PropertyTag ContainerClassW = new PropertyTag(PropertyId.ContainerClass, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CONTAINER_CONTENTS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTAINER_CONTENTS.
        /// </remarks>
        public static readonly PropertyTag ContainerContents = new PropertyTag(PropertyId.ContainerContents, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_CONTAINER_FLAGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTAINER_FLAGS.
        /// </remarks>
        public static readonly PropertyTag ContainerFlags = new PropertyTag(PropertyId.ContainerFlags, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_CONTAINER_HIERARCHY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTAINER_HIERARCHY.
        /// </remarks>
        public static readonly PropertyTag ContainerHierarchy = new PropertyTag(PropertyId.ContainerHierarchy, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_CONTAINER_MODIFY_VERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTAINER_MODIFY_VERSION.
        /// </remarks>
        public static readonly PropertyTag ContainerModifyVersion = new PropertyTag(PropertyId.ContainerModifyVersion, PropertyType.I8);

        /// <summary>
        /// The MAPI property PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID.
        /// </remarks>
        public static readonly PropertyTag ContentConfidentialityAlgorithmId = new PropertyTag(PropertyId.ContentConfidentialityAlgorithmId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONTENT_CORRELATOR.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_CORRELATOR.
        /// </remarks>
        public static readonly PropertyTag ContentCorrelator = new PropertyTag(PropertyId.ContentCorrelator, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONTENT_COUNT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_COUNT.
        /// </remarks>
        public static readonly PropertyTag ContentCount = new PropertyTag(PropertyId.ContentCount, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_CONTENT_IDENTIFIER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_IDENTIFIER.
        /// </remarks>
        public static readonly PropertyTag ContentIdentifierA = new PropertyTag(PropertyId.ContentIdentifier, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_CONTENT_IDENTIFIER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_IDENTIFIER.
        /// </remarks>
        public static readonly PropertyTag ContentIdentifierW = new PropertyTag(PropertyId.ContentIdentifier, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CONTENT_INTEGRITY_CHECK.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_INTEGRITY_CHECK.
        /// </remarks>
        public static readonly PropertyTag ContentIntegrityCheck = new PropertyTag(PropertyId.ContentIntegrityCheck, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONTENT_LENGTH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_LENGTH.
        /// </remarks>
        public static readonly PropertyTag ContentLength = new PropertyTag(PropertyId.ContentLength, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_CONTENT_RETURN_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_RETURN_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag ContentReturnRequested = new PropertyTag(PropertyId.ContentReturnRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_CONTENTS_SORT_ORDER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENTS_SORT_ORDER.
        /// </remarks>
        public static readonly PropertyTag ContentsSortOrder = new PropertyTag(PropertyId.ContentsSortOrder, PropertyType.Long, true);

        /// <summary>
        /// The MAPI property PR_CONTENT_UNREAD.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTENT_UNREAD.
        /// </remarks>
        public static readonly PropertyTag ContentUnread = new PropertyTag(PropertyId.ContentUnread, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_CONTROL_FLAGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTROL_FLAGS.
        /// </remarks>
        public static readonly PropertyTag ControlFlags = new PropertyTag(PropertyId.ControlFlags, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_CONTROL_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTROL_ID.
        /// </remarks>
        public static readonly PropertyTag ControlId = new PropertyTag(PropertyId.ControlId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONTROL_STRUCTURE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTROL_STRUCTURE.
        /// </remarks>
        public static readonly PropertyTag ControlStructure = new PropertyTag(PropertyId.ControlStructure, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONTROL_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONTROL_TYPE.
        /// </remarks>
        public static readonly PropertyTag ControlType = new PropertyTag(PropertyId.ControlType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_CONVERSATION_INDEX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERSATION_INDEX.
        /// </remarks>
        public static readonly PropertyTag ConversationIndex = new PropertyTag(PropertyId.ConversationIndex, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONVERSATION_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERSATION_KEY.
        /// </remarks>
        public static readonly PropertyTag ConversationKey = new PropertyTag(PropertyId.ConversationKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONVERSATION_TOPIC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERSATION_TOPIC.
        /// </remarks>
        public static readonly PropertyTag ConversationTopicA = new PropertyTag(PropertyId.ConversationTopic, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_CONVERSATION_TOPIC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERSATION_TOPIC.
        /// </remarks>
        public static readonly PropertyTag ConversationTopicW = new PropertyTag(PropertyId.ConversationTopic, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CONVERSION_EITS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERSION_EITS.
        /// </remarks>
        public static readonly PropertyTag ConversionEits = new PropertyTag(PropertyId.ConversionEits, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CONVERSION_PROHIBITED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERSION_PROHIBITED.
        /// </remarks>
        public static readonly PropertyTag ConversionProhibited = new PropertyTag(PropertyId.ConversionProhibited, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_CONVERSION_WITH_LOSS_PROHIBITED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERSION_WITH_LOSS_PROHIBITED.
        /// </remarks>
        public static readonly PropertyTag ConversionWithLossProhibited = new PropertyTag(PropertyId.ConversionWithLossProhibited, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_CONVERTED_EITS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CONVERTED_EITS.
        /// </remarks>
        public static readonly PropertyTag ConvertedEits = new PropertyTag(PropertyId.ConvertedEits, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_CORRELATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CORRELATE.
        /// </remarks>
        public static readonly PropertyTag Correlate = new PropertyTag(PropertyId.Correlate, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_CORRELATE_MTSID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CORRELATE_MTSID.
        /// </remarks>
        public static readonly PropertyTag CorrelateMtsid = new PropertyTag(PropertyId.CorrelateMtsid, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag CountryA = new PropertyTag(PropertyId.Country, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag CountryW = new PropertyTag(PropertyId.Country, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_CREATE_TEMPLATES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CREATE_TEMPLATES.
        /// </remarks>
        public static readonly PropertyTag CreateTemplates = new PropertyTag(PropertyId.CreateTemplates, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_CREATION_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CREATION_TIME.
        /// </remarks>
        public static readonly PropertyTag CreationTime = new PropertyTag(PropertyId.CreationTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_CREATION_VERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CREATION_VERSION.
        /// </remarks>
        public static readonly PropertyTag CreationVersion = new PropertyTag(PropertyId.CreationVersion, PropertyType.I8);

        /// <summary>
        /// The MAPI property PR_CURRENT_VERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CURRENT_VERSION.
        /// </remarks>
        public static readonly PropertyTag CurrentVersion = new PropertyTag(PropertyId.CurrentVersion, PropertyType.I8);

        /// <summary>
        /// The MAPI property PR_CUSTOMER_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CUSTOMER_ID.
        /// </remarks>
        public static readonly PropertyTag CustomerIdA = new PropertyTag(PropertyId.CustomerId, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_CUSTOMER_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_CUSTOMER_ID.
        /// </remarks>
        public static readonly PropertyTag CustomerIdW = new PropertyTag(PropertyId.CustomerId, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_DEFAULT_PROFILE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEFAULT_PROFILE.
        /// </remarks>
        public static readonly PropertyTag DefaultProfile = new PropertyTag(PropertyId.DefaultProfile, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_DEFAULT_STORE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEFAULT_STORE.
        /// </remarks>
        public static readonly PropertyTag DefaultStore = new PropertyTag(PropertyId.DefaultStore, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_DEFAULT_VIEW_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEFAULT_VIEW_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag DefaultViewEntryId = new PropertyTag(PropertyId.DefaultViewEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_DEF_CREATE_DL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEF_CREATE_DL.
        /// </remarks>
        public static readonly PropertyTag DefCreateDl = new PropertyTag(PropertyId.DefCreateDl, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_DEF_CREATE_MAILUSER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEF_CREATE_MAILUSER.
        /// </remarks>
        public static readonly PropertyTag DefCreateMailuser = new PropertyTag(PropertyId.DefCreateMailuser, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_DEFERRED_DELIVERY_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEFERRED_DELIVERY_TIME.
        /// </remarks>
        public static readonly PropertyTag DeferredDeliveryTime = new PropertyTag(PropertyId.DeferredDeliveryTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_DELEGATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DELEGATION.
        /// </remarks>
        public static readonly PropertyTag Delegation = new PropertyTag(PropertyId.Delegation, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_DELETE_AFTER_SUBMIT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DELETE_AFTER_SUBMIT.
        /// </remarks>
        public static readonly PropertyTag DeleteAfterSubmit = new PropertyTag(PropertyId.DeleteAfterSubmit, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_DELIVER_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DELIVER_TIME.
        /// </remarks>
        public static readonly PropertyTag DeliverTime = new PropertyTag(PropertyId.DeliverTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_DELIVERY_POINT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DELIVERY_POINT.
        /// </remarks>
        public static readonly PropertyTag DeliveryPoint = new PropertyTag(PropertyId.DeliveryPoint, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_DELTAX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DELTAX.
        /// </remarks>
        public static readonly PropertyTag Deltax = new PropertyTag(PropertyId.Deltax, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_DELTAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DELTAY.
        /// </remarks>
        public static readonly PropertyTag Deltay = new PropertyTag(PropertyId.Deltay, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_DEPARTMENT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEPARTMENT_NAME.
        /// </remarks>
        public static readonly PropertyTag DepartmentNameA = new PropertyTag(PropertyId.DepartmentName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_DEPARTMENT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEPARTMENT_NAME.
        /// </remarks>
        public static readonly PropertyTag DepartmentNameW = new PropertyTag(PropertyId.DepartmentName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_DEPTH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DEPTH.
        /// </remarks>
        public static readonly PropertyTag Depth = new PropertyTag(PropertyId.Depth, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_DETAILS_TABLE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DETAILS_TABLE.
        /// </remarks>
        public static readonly PropertyTag DetailsTable = new PropertyTag(PropertyId.DetailsTable, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_DISCARD_REASON.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISCARD_REASON.
        /// </remarks>
        public static readonly PropertyTag DiscardReason = new PropertyTag(PropertyId.DiscardReason, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_DISCLOSE_RECIPIENTS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISCLOSE_RECIPIENTS.
        /// </remarks>
        public static readonly PropertyTag DiscloseRecipients = new PropertyTag(PropertyId.DiscloseRecipients, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_DISCLOSURE_OF_RECIPIENTS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISCLOSURE_OF_RECIPIENTS.
        /// </remarks>
        public static readonly PropertyTag DisclosureOfRecipients = new PropertyTag(PropertyId.DisclosureOfRecipients, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_DISCRETE_VALUES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISCRETE_VALUES.
        /// </remarks>
        public static readonly PropertyTag DiscreteValues = new PropertyTag(PropertyId.DiscreteValues, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_DISC_VAL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISC_VAL.
        /// </remarks>
        public static readonly PropertyTag DiscVal = new PropertyTag(PropertyId.DiscVal, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_DISPLAY_BCC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_BCC.
        /// </remarks>
        public static readonly PropertyTag DisplayBccA = new PropertyTag(PropertyId.DisplayBcc, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_DISPLAY_BCC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_BCC.
        /// </remarks>
        public static readonly PropertyTag DisplayBccW = new PropertyTag(PropertyId.DisplayBcc, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_DISPLAY_CC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_CC.
        /// </remarks>
        public static readonly PropertyTag DisplayCcA = new PropertyTag(PropertyId.DisplayCc, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_DISPLAY_CC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_CC.
        /// </remarks>
        public static readonly PropertyTag DisplayCcW = new PropertyTag(PropertyId.DisplayCc, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag DisplayNameA = new PropertyTag(PropertyId.DisplayName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag DisplayNameW = new PropertyTag(PropertyId.DisplayName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_DISPLAY_NAME_PREFIX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_NAME_PREFIX.
        /// </remarks>
        public static readonly PropertyTag DisplayNamePrefixA = new PropertyTag(PropertyId.DisplayNamePrefix, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_DISPLAY_NAME_PREFIX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_NAME_PREFIX.
        /// </remarks>
        public static readonly PropertyTag DisplayNamePrefixW = new PropertyTag(PropertyId.DisplayNamePrefix, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_DISPLAY_TO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_TO.
        /// </remarks>
        public static readonly PropertyTag DisplayToA = new PropertyTag(PropertyId.DisplayTo, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_DISPLAY_TO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_TO.
        /// </remarks>
        public static readonly PropertyTag DisplayToW = new PropertyTag(PropertyId.DisplayTo, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_DISPLAY_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DISPLAY_TYPE.
        /// </remarks>
        public static readonly PropertyTag DisplayType = new PropertyTag(PropertyId.DisplayType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_DL_EXPANSION_HISTORY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DL_EXPANSION_HISTORY.
        /// </remarks>
        public static readonly PropertyTag DlExpansionHistory = new PropertyTag(PropertyId.DlExpansionHistory, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_DL_EXPANSION_PROHIBITED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_DL_EXPANSION_PROHIBITED.
        /// </remarks>
        public static readonly PropertyTag DlExpansionProhibited = new PropertyTag(PropertyId.DlExpansionProhibited, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag EmailAddressA = new PropertyTag(PropertyId.EmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag EmailAddressW = new PropertyTag(PropertyId.EmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_END_DATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_END_DATE.
        /// </remarks>
        public static readonly PropertyTag EndDate = new PropertyTag(PropertyId.EndDate, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag EntryId = new PropertyTag(PropertyId.EntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_EXPAND_BEGIN_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EXPAND_BEGIN_TIME.
        /// </remarks>
        public static readonly PropertyTag ExpandBeginTime = new PropertyTag(PropertyId.ExpandBeginTime, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_EXPANDED_BEGIN_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EXPANDED_BEGIN_TIME.
        /// </remarks>
        public static readonly PropertyTag ExpandedBeginTime = new PropertyTag(PropertyId.ExpandedBeginTime, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_EXPANDED_END_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EXPANDED_END_TIME.
        /// </remarks>
        public static readonly PropertyTag ExpandedEndTime = new PropertyTag(PropertyId.ExpandedEndTime, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_EXPAND_END_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EXPAND_END_TIME.
        /// </remarks>
        public static readonly PropertyTag ExpandEndTime = new PropertyTag(PropertyId.ExpandEndTime, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_EXPIRY_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EXPIRY_TIME.
        /// </remarks>
        public static readonly PropertyTag ExpiryTime = new PropertyTag(PropertyId.ExpiryTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_EXPLICIT_CONVERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_EXPLICIT_CONVERSION.
        /// </remarks>
        public static readonly PropertyTag ExplicitConversion = new PropertyTag(PropertyId.ExplicitConversion, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_FILTERING_HOOKS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FILTERING_HOOKS.
        /// </remarks>
        public static readonly PropertyTag FilteringHooks = new PropertyTag(PropertyId.FilteringHooks, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_FINDER_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FINDER_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag FinderEntryId = new PropertyTag(PropertyId.FinderEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_FOLDER_ASSOCIATED_CONTENTS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FOLDER_ASSOCIATED_CONTENTS.
        /// </remarks>
        public static readonly PropertyTag FolderAssociatedContents = new PropertyTag(PropertyId.FolderAssociatedContents, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_FOLDER_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FOLDER_TYPE.
        /// </remarks>
        public static readonly PropertyTag FolderType = new PropertyTag(PropertyId.FolderType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_FORM_CATEGORY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_CATEGORY.
        /// </remarks>
        public static readonly PropertyTag FormCategoryA = new PropertyTag(PropertyId.FormCategory, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_FORM_CATEGORY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_CATEGORY.
        /// </remarks>
        public static readonly PropertyTag FormCategoryW = new PropertyTag(PropertyId.FormCategory, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_FORM_CATEGORY_SUB.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_CATEGORY_SUB.
        /// </remarks>
        public static readonly PropertyTag FormCategorySubA = new PropertyTag(PropertyId.FormCategorySub, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_FORM_CATEGORY_SUB.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_CATEGORY_SUB.
        /// </remarks>
        public static readonly PropertyTag FormCategorySubW = new PropertyTag(PropertyId.FormCategorySub, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_FORM_CLSID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_CLSID.
        /// </remarks>
        public static readonly PropertyTag FormClsid = new PropertyTag(PropertyId.FormClsid, PropertyType.ClassId);

        /// <summary>
        /// The MAPI property PR_FORM_CONTACT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_CONTACT_NAME.
        /// </remarks>
        public static readonly PropertyTag FormContactNameA = new PropertyTag(PropertyId.FormContactName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_FORM_CONTACT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_CONTACT_NAME.
        /// </remarks>
        public static readonly PropertyTag FormContactNameW = new PropertyTag(PropertyId.FormContactName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_FORM_DESIGNER_GUID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_DESIGNER_GUID.
        /// </remarks>
        public static readonly PropertyTag FormDesignerGuid = new PropertyTag(PropertyId.FormDesignerGuid, PropertyType.ClassId);

        /// <summary>
        /// The MAPI property PR_FORM_DESIGNER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_DESIGNER_NAME.
        /// </remarks>
        public static readonly PropertyTag FormDesignerNameA = new PropertyTag(PropertyId.FormDesignerName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_FORM_DESIGNER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_DESIGNER_NAME.
        /// </remarks>
        public static readonly PropertyTag FormDesignerNameW = new PropertyTag(PropertyId.FormDesignerName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_FORM_HIDDEN.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_HIDDEN.
        /// </remarks>
        public static readonly PropertyTag FormHidden = new PropertyTag(PropertyId.FormHidden, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_FORM_HOST_MAP.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_HOST_MAP.
        /// </remarks>
        public static readonly PropertyTag FormHostMap = new PropertyTag(PropertyId.FormHostMap, PropertyType.Long, true);

        /// <summary>
        /// The MAPI property PR_FORM_MESSAGE_BEHAVIOR.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_MESSAGE_BEHAVIOR.
        /// </remarks>
        public static readonly PropertyTag FormMessageBehavior = new PropertyTag(PropertyId.FormMessageBehavior, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_FORM_VERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_VERSION.
        /// </remarks>
        public static readonly PropertyTag FormVersionA = new PropertyTag(PropertyId.FormVersion, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_FORM_VERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FORM_VERSION.
        /// </remarks>
        public static readonly PropertyTag FormVersionW = new PropertyTag(PropertyId.FormVersion, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_FTP_SITE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FTP_SITE.
        /// </remarks>
        public static readonly PropertyTag FtpSiteA = new PropertyTag(PropertyId.FtpSite, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_FTP_SITE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_FTP_SITE.
        /// </remarks>
        public static readonly PropertyTag FtpSiteW = new PropertyTag(PropertyId.FtpSite, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_GENDER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_GENDER.
        /// </remarks>
        public static readonly PropertyTag Gender = new PropertyTag(PropertyId.Gender, PropertyType.I2);

        /// <summary>
        /// The MAPI property PR_GENERATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_GENERATION.
        /// </remarks>
        public static readonly PropertyTag GenerationA = new PropertyTag(PropertyId.Generation, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_GENERATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_GENERATION.
        /// </remarks>
        public static readonly PropertyTag GenerationW = new PropertyTag(PropertyId.Generation, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_GIVEN_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_GIVEN_NAME.
        /// </remarks>
        public static readonly PropertyTag GivenNameA = new PropertyTag(PropertyId.GivenName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_GIVEN_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_GIVEN_NAME.
        /// </remarks>
        public static readonly PropertyTag GivenNameW = new PropertyTag(PropertyId.GivenName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_GOVERNMENT_ID_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_GOVERNMENT_ID_NUMBER.
        /// </remarks>
        public static readonly PropertyTag GovernmentIdNumberA = new PropertyTag(PropertyId.GovernmentIdNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_GOVERNMENT_ID_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_GOVERNMENT_ID_NUMBER.
        /// </remarks>
        public static readonly PropertyTag GovernmentIdNumberW = new PropertyTag(PropertyId.GovernmentIdNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HASATTACH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HASATTACH.
        /// </remarks>
        public static readonly PropertyTag Hasattach = new PropertyTag(PropertyId.Hasattach, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_HEADER_FOLDER_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HEADER_FOLDER_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag HeaderFolderEntryId = new PropertyTag(PropertyId.HeaderFolderEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_HOBBIES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOBBIES.
        /// </remarks>
        public static readonly PropertyTag HobbiesA = new PropertyTag(PropertyId.Hobbies, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOBBIES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOBBIES.
        /// </remarks>
        public static readonly PropertyTag HobbiesW = new PropertyTag(PropertyId.Hobbies, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Home2TelephoneNumberA = new PropertyTag(PropertyId.Home2TelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Home2TelephoneNumberAMv = new PropertyTag(PropertyId.Home2TelephoneNumber, PropertyType.String8, true);

        /// <summary>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Home2TelephoneNumberW = new PropertyTag(PropertyId.Home2TelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Home2TelephoneNumberWMv = new PropertyTag(PropertyId.Home2TelephoneNumber, PropertyType.Unicode, true);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_CITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_CITY.
        /// </remarks>
        public static readonly PropertyTag HomeAddressCityA = new PropertyTag(PropertyId.HomeAddressCity, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_CITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_CITY.
        /// </remarks>
        public static readonly PropertyTag HomeAddressCityW = new PropertyTag(PropertyId.HomeAddressCity, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag HomeAddressCountryA = new PropertyTag(PropertyId.HomeAddressCountry, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag HomeAddressCountryW = new PropertyTag(PropertyId.HomeAddressCountry, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag HomeAddressPostalCodeA = new PropertyTag(PropertyId.HomeAddressPostalCode, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag HomeAddressPostalCodeW = new PropertyTag(PropertyId.HomeAddressPostalCode, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
        /// </remarks>
        public static readonly PropertyTag HomeAddressPostOfficeBoxA = new PropertyTag(PropertyId.HomeAddressPostOfficeBox, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
        /// </remarks>
        public static readonly PropertyTag HomeAddressPostOfficeBoxW = new PropertyTag(PropertyId.HomeAddressPostOfficeBox, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
        /// </remarks>
        public static readonly PropertyTag HomeAddressStateOrProvinceA = new PropertyTag(PropertyId.HomeAddressStateOrProvince, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
        /// </remarks>
        public static readonly PropertyTag HomeAddressStateOrProvinceW = new PropertyTag(PropertyId.HomeAddressStateOrProvince, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_STREET.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_STREET.
        /// </remarks>
        public static readonly PropertyTag HomeAddressStreetA = new PropertyTag(PropertyId.HomeAddressStreet, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_ADDRESS_STREET.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_ADDRESS_STREET.
        /// </remarks>
        public static readonly PropertyTag HomeAddressStreetW = new PropertyTag(PropertyId.HomeAddressStreet, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME_FAX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_FAX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag HomeFaxNumberA = new PropertyTag(PropertyId.HomeFaxNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_FAX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_FAX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag HomeFaxNumberW = new PropertyTag(PropertyId.HomeFaxNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_HOME_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag HomeTelephoneNumberA = new PropertyTag(PropertyId.HomeTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_HOME_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_HOME_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag HomeTelephoneNumberW = new PropertyTag(PropertyId.HomeTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ICON.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ICON.
        /// </remarks>
        public static readonly PropertyTag Icon = new PropertyTag(PropertyId.Icon, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IDENTITY_DISPLAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IDENTITY_DISPLAY.
        /// </remarks>
        public static readonly PropertyTag IdentityDisplayA = new PropertyTag(PropertyId.IdentityDisplay, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_IDENTITY_DISPLAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IDENTITY_DISPLAY.
        /// </remarks>
        public static readonly PropertyTag IdentityDisplayW = new PropertyTag(PropertyId.IdentityDisplay, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_IDENTITY_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IDENTITY_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag IdentityEntryId = new PropertyTag(PropertyId.IdentityEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IDENTITY_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IDENTITY_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag IdentitySearchKey = new PropertyTag(PropertyId.IdentitySearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IMPLICIT_CONVERSION_PROHIBITED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IMPLICIT_CONVERSION_PROHIBITED.
        /// </remarks>
        public static readonly PropertyTag ImplicitConversionProhibited = new PropertyTag(PropertyId.ImplicitConversionProhibited, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_IMPORTANCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IMPORTANCE.
        /// </remarks>
        public static readonly PropertyTag Importance = new PropertyTag(PropertyId.Importance, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_INCOMPLETE_COPY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INCOMPLETE_COPY.
        /// </remarks>
        public static readonly PropertyTag IncompleteCopy = new PropertyTag(PropertyId.IncompleteCopy, PropertyType.Boolean);

        /// <summary>
        /// The Internet mail override charset.
        /// </summary>
        /// <remarks>
        /// The Internet mail override charset.
        /// </remarks>
        public static readonly PropertyTag INetMailOverrideCharset = new PropertyTag(PropertyId.INetMailOverrideCharset, PropertyType.Unspecified);

        /// <summary>
        /// The Internet mail override format.
        /// </summary>
        /// <remarks>
        /// The Internet mail override format.
        /// </remarks>
        public static readonly PropertyTag INetMailOverrideFormat = new PropertyTag(PropertyId.INetMailOverrideFormat, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_INITIAL_DETAILS_PANE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INITIAL_DETAILS_PANE.
        /// </remarks>
        public static readonly PropertyTag InitialDetailsPane = new PropertyTag(PropertyId.InitialDetailsPane, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_INITIALS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INITIALS.
        /// </remarks>
        public static readonly PropertyTag InitialsA = new PropertyTag(PropertyId.Initials, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INITIALS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INITIALS.
        /// </remarks>
        public static readonly PropertyTag InitialsW = new PropertyTag(PropertyId.Initials, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_IN_REPLY_TO_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IN_REPLY_TO_ID.
        /// </remarks>
        public static readonly PropertyTag InReplyToIdA = new PropertyTag(PropertyId.InReplyToId, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_IN_REPLY_TO_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IN_REPLY_TO_ID.
        /// </remarks>
        public static readonly PropertyTag InReplyToIdW = new PropertyTag(PropertyId.InReplyToId, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INSTANCE_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INSTANCE_KEY.
        /// </remarks>
        public static readonly PropertyTag InstanceKey = new PropertyTag(PropertyId.InstanceKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_INTERNET_APPROVED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_APPROVED.
        /// </remarks>
        public static readonly PropertyTag InternetApprovedA = new PropertyTag(PropertyId.InternetApproved, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_APPROVED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_APPROVED.
        /// </remarks>
        public static readonly PropertyTag InternetApprovedW = new PropertyTag(PropertyId.InternetApproved, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_ARTICLE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_ARTICLE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag InternetArticleNumber = new PropertyTag(PropertyId.InternetArticleNumber, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_INTERNET_CONTROL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_CONTROL.
        /// </remarks>
        public static readonly PropertyTag InternetControlA = new PropertyTag(PropertyId.InternetControl, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_CONTROL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_CONTROL.
        /// </remarks>
        public static readonly PropertyTag InternetControlW = new PropertyTag(PropertyId.InternetControl, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_CPID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_CPID.
        /// </remarks>
        public static readonly PropertyTag InternetCPID = new PropertyTag(PropertyId.InternetCPID, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_INTERNET_DISTRIBUTION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_DISTRIBUTION.
        /// </remarks>
        public static readonly PropertyTag InternetDistributionA = new PropertyTag(PropertyId.InternetDistribution, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_DISTRIBUTION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_DISTRIBUTION.
        /// </remarks>
        public static readonly PropertyTag InternetDistributionW = new PropertyTag(PropertyId.InternetDistribution, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_FOLLOWUP_TO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_FOLLOWUP_TO.
        /// </remarks>
        public static readonly PropertyTag InternetFollowupToA = new PropertyTag(PropertyId.InternetFollowupTo, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_FOLLOWUP_TO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_FOLLOWUP_TO.
        /// </remarks>
        public static readonly PropertyTag InternetFollowupToW = new PropertyTag(PropertyId.InternetFollowupTo, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_LINES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_LINES.
        /// </remarks>
        public static readonly PropertyTag InternetLines = new PropertyTag(PropertyId.InternetLines, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_INTERNET_MESSAGE_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_MESSAGE_ID.
        /// </remarks>
        public static readonly PropertyTag InternetMessageIdA = new PropertyTag(PropertyId.InternetMessageId, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_MESSAGE_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_MESSAGE_ID.
        /// </remarks>
        public static readonly PropertyTag InternetMessageIdW = new PropertyTag(PropertyId.InternetMessageId, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_NEWSGROUPS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_NEWSGROUPS.
        /// </remarks>
        public static readonly PropertyTag InternetNewsgroupsA = new PropertyTag(PropertyId.InternetNewsgroups, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_NEWSGROUPS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_NEWSGROUPS.
        /// </remarks>
        public static readonly PropertyTag InternetNewsgroupsW = new PropertyTag(PropertyId.InternetNewsgroups, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_NNTP_PATH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_NNTP_PATH.
        /// </remarks>
        public static readonly PropertyTag InternetNntpPathA = new PropertyTag(PropertyId.InternetNntpPath, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_NNTP_PATH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_NNTP_PATH.
        /// </remarks>
        public static readonly PropertyTag InternetNntpPathW = new PropertyTag(PropertyId.InternetNntpPath, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_ORGANIZATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_ORGANIZATION.
        /// </remarks>
        public static readonly PropertyTag InternetOrganizationA = new PropertyTag(PropertyId.InternetOrganization, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_ORGANIZATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_ORGANIZATION.
        /// </remarks>
        public static readonly PropertyTag InternetOrganizationW = new PropertyTag(PropertyId.InternetOrganization, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_PRECEDENCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_PRECEDENCE.
        /// </remarks>
        public static readonly PropertyTag InternetPrecedenceA = new PropertyTag(PropertyId.InternetPrecedence, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_PRECEDENCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_PRECEDENCE.
        /// </remarks>
        public static readonly PropertyTag InternetPrecedenceW = new PropertyTag(PropertyId.InternetPrecedence, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_INTERNET_REFERENCES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_REFERENCES.
        /// </remarks>
        public static readonly PropertyTag InternetReferencesA = new PropertyTag(PropertyId.InternetReferences, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_INTERNET_REFERENCES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_INTERNET_REFERENCES.
        /// </remarks>
        public static readonly PropertyTag InternetReferencesW = new PropertyTag(PropertyId.InternetReferences, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_IPM_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_ID.
        /// </remarks>
        public static readonly PropertyTag IpmId = new PropertyTag(PropertyId.IpmId, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_IPM_OUTBOX_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_OUTBOX_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag IpmOutboxEntryId = new PropertyTag(PropertyId.IpmOutboxEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IPM_OUTBOX_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_OUTBOX_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag IpmOutboxSearchKey = new PropertyTag(PropertyId.IpmOutboxSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IPM_RETURN_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_RETURN_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag IpmReturnRequested = new PropertyTag(PropertyId.IpmReturnRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_IPM_SENTMAIL_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_SENTMAIL_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag IpmSentmailEntryId = new PropertyTag(PropertyId.IpmSentmailEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IPM_SENTMAIL_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_SENTMAIL_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag IpmSentmailSearchKey = new PropertyTag(PropertyId.IpmSentmailSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IPM_SUBTREE_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_SUBTREE_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag IpmSubtreeEntryId = new PropertyTag(PropertyId.IpmSubtreeEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IPM_SUBTREE_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_SUBTREE_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag IpmSubtreeSearchKey = new PropertyTag(PropertyId.IpmSubtreeSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IPM_WASTEBASKET_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_WASTEBASKET_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag IpmWastebasketEntryId = new PropertyTag(PropertyId.IpmWastebasketEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_IPM_WASTEBASKET_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_IPM_WASTEBASKET_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag IpmWastebasketSearchKey = new PropertyTag(PropertyId.IpmWastebasketSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ISDN_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ISDN_NUMBER.
        /// </remarks>
        public static readonly PropertyTag IsdnNumberA = new PropertyTag(PropertyId.IsdnNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ISDN_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ISDN_NUMBER.
        /// </remarks>
        public static readonly PropertyTag IsdnNumberW = new PropertyTag(PropertyId.IsdnNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_KEYWORD.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_KEYWORD.
        /// </remarks>
        public static readonly PropertyTag KeywordA = new PropertyTag(PropertyId.Keyword, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_KEYWORD.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_KEYWORD.
        /// </remarks>
        public static readonly PropertyTag KeywordW = new PropertyTag(PropertyId.Keyword, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LANGUAGE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LANGUAGE.
        /// </remarks>
        public static readonly PropertyTag LanguageA = new PropertyTag(PropertyId.Language, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LANGUAGE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LANGUAGE.
        /// </remarks>
        public static readonly PropertyTag LanguageW = new PropertyTag(PropertyId.Language, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LANGUAGES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LANGUAGES.
        /// </remarks>
        public static readonly PropertyTag LanguagesA = new PropertyTag(PropertyId.Languages, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LANGUAGES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LANGUAGES.
        /// </remarks>
        public static readonly PropertyTag LanguagesW = new PropertyTag(PropertyId.Languages, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LAST_MODIFICATION_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LAST_MODIFICATION_TIME.
        /// </remarks>
        public static readonly PropertyTag LastModificationTime = new PropertyTag(PropertyId.LastModificationTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_LAST_MODIFIER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LAST_MODIFIER_NAME.
        /// </remarks>
        public static readonly PropertyTag LastModifierNameA = new PropertyTag(PropertyId.LastModifierName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LAST_MODIFIER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LAST_MODIFIER_NAME.
        /// </remarks>
        public static readonly PropertyTag LastModifierNameW = new PropertyTag(PropertyId.LastModifierName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LATEST_DELIVERY_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LATEST_DELIVERY_TIME.
        /// </remarks>
        public static readonly PropertyTag LatestDeliveryTime = new PropertyTag(PropertyId.LatestDeliveryTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_LIST_HELP.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LIST_HELP.
        /// </remarks>
        public static readonly PropertyTag ListHelpA = new PropertyTag(PropertyId.ListHelp, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LIST_HELP.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LIST_HELP.
        /// </remarks>
        public static readonly PropertyTag ListHelpW = new PropertyTag(PropertyId.ListHelp, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LIST_SUBSCRIBE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LIST_SUBSCRIBE.
        /// </remarks>
        public static readonly PropertyTag ListSubscribeA = new PropertyTag(PropertyId.ListSubscribe, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LIST_SUBSCRIBE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LIST_SUBSCRIBE.
        /// </remarks>
        public static readonly PropertyTag ListSubscribeW = new PropertyTag(PropertyId.ListSubscribe, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LIST_UNSUBSCRIBE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LIST_UNSUBSCRIBE.
        /// </remarks>
        public static readonly PropertyTag ListUnsubscribeA = new PropertyTag(PropertyId.ListUnsubscribe, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LIST_UNSUBSCRIBE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LIST_UNSUBSCRIBE.
        /// </remarks>
        public static readonly PropertyTag ListUnsubscribeW = new PropertyTag(PropertyId.ListUnsubscribe, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LOCALITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCALITY.
        /// </remarks>
        public static readonly PropertyTag LocalityA = new PropertyTag(PropertyId.Locality, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LOCALITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCALITY.
        /// </remarks>
        public static readonly PropertyTag LocalityW = new PropertyTag(PropertyId.Locality, PropertyType.Unicode);

        //public static readonly PropertyTag LocallyDelivered = new PropertyTag (TnefPropertyId.LocallyDelivered, TnefPropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCATION.
        /// </remarks>
        public static readonly PropertyTag LocationA = new PropertyTag(PropertyId.Location, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCATION.
        /// </remarks>
        public static readonly PropertyTag LocationW = new PropertyTag(PropertyId.Location, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_LOCK_BRANCH_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_BRANCH_ID.
        /// </remarks>
        public static readonly PropertyTag LockBranchId = new PropertyTag(PropertyId.LockBranchId, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_DEPTH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_DEPTH.
        /// </remarks>
        public static readonly PropertyTag LockDepth = new PropertyTag(PropertyId.LockDepth, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_ENLISTMENT_CONTEXT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_ENLISTMENT_CONTEXT.
        /// </remarks>
        public static readonly PropertyTag LockEnlistmentContext = new PropertyTag(PropertyId.LockEnlistmentContext, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_EXPIRY_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_EXPIRY_TIME.
        /// </remarks>
        public static readonly PropertyTag LockExpiryTime = new PropertyTag(PropertyId.LockExpiryTime, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_PERSISTENT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_PERSISTENT.
        /// </remarks>
        public static readonly PropertyTag LockPersistent = new PropertyTag(PropertyId.LockPersistent, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_RESOURCE_DID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_RESOURCE_DID.
        /// </remarks>
        public static readonly PropertyTag LockResourceDid = new PropertyTag(PropertyId.LockResourceDid, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_RESOURCE_FID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_RESOURCE_FID.
        /// </remarks>
        public static readonly PropertyTag LockResourceFid = new PropertyTag(PropertyId.LockResourceFid, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_RESOURCE_MID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_RESOURCE_MID.
        /// </remarks>
        public static readonly PropertyTag LockResourceMid = new PropertyTag(PropertyId.LockResourceMid, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_SCOPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_SCOPE.
        /// </remarks>
        public static readonly PropertyTag LockScope = new PropertyTag(PropertyId.LockScope, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_TIMEOUT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_TIMEOUT.
        /// </remarks>
        public static readonly PropertyTag LockTimeout = new PropertyTag(PropertyId.LockTimeout, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_LOCK_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_LOCK_TYPE.
        /// </remarks>
        public static readonly PropertyTag LockType = new PropertyTag(PropertyId.LockType, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_MAIL_PERMISSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MAIL_PERMISSION.
        /// </remarks>
        public static readonly PropertyTag MailPermission = new PropertyTag(PropertyId.MailPermission, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_MANAGER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MANAGER_NAME.
        /// </remarks>
        public static readonly PropertyTag ManagerNameA = new PropertyTag(PropertyId.ManagerName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_MANAGER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MANAGER_NAME.
        /// </remarks>
        public static readonly PropertyTag ManagerNameW = new PropertyTag(PropertyId.ManagerName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_MAPPING_SIGNATURE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MAPPING_SIGNATURE.
        /// </remarks>
        public static readonly PropertyTag MappingSignature = new PropertyTag(PropertyId.MappingSignature, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_MDB_PROVIDER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MDB_PROVIDER.
        /// </remarks>
        public static readonly PropertyTag MdbProvider = new PropertyTag(PropertyId.MdbProvider, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_MESSAGE_ATTACHMENTS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_ATTACHMENTS.
        /// </remarks>
        public static readonly PropertyTag MessageAttachments = new PropertyTag(PropertyId.MessageAttachments, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_MESSAGE_CC_ME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_CC_ME.
        /// </remarks>
        public static readonly PropertyTag MessageCcMe = new PropertyTag(PropertyId.MessageCcMe, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_MESSAGE_CLASS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_CLASS.
        /// </remarks>
        public static readonly PropertyTag MessageClassA = new PropertyTag(PropertyId.MessageClass, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_MESSAGE_CLASS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_CLASS.
        /// </remarks>
        public static readonly PropertyTag MessageClassW = new PropertyTag(PropertyId.MessageClass, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_MESSAGE_CODEPAGE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_CODEPAGE.
        /// </remarks>
        public static readonly PropertyTag MessageCodepage = new PropertyTag(PropertyId.MessageCodepage, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_MESSAGE_DELIVERY_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_DELIVERY_ID.
        /// </remarks>
        public static readonly PropertyTag MessageDeliveryId = new PropertyTag(PropertyId.MessageDeliveryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_MESSAGE_DELIVERY_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_DELIVERY_TIME.
        /// </remarks>
        public static readonly PropertyTag MessageDeliveryTime = new PropertyTag(PropertyId.MessageDeliveryTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_MESSAGE_DOWNLOAD_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_DOWNLOAD_TIME.
        /// </remarks>
        public static readonly PropertyTag MessageDownloadTime = new PropertyTag(PropertyId.MessageDownloadTime, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_MESSAGE_FLAGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_FLAGS.
        /// </remarks>
        public static readonly PropertyTag MessageFlags = new PropertyTag(PropertyId.MessageFlags, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_MESSAGE_RECIPIENTS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_RECIPIENTS.
        /// </remarks>
        public static readonly PropertyTag MessageRecipients = new PropertyTag(PropertyId.MessageRecipients, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_MESSAGE_RECIP_ME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_RECIP_ME.
        /// </remarks>
        public static readonly PropertyTag MessageRecipMe = new PropertyTag(PropertyId.MessageRecipMe, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_MESSAGE_SECURITY_LABEL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_SECURITY_LABEL.
        /// </remarks>
        public static readonly PropertyTag MessageSecurityLabel = new PropertyTag(PropertyId.MessageSecurityLabel, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_MESSAGE_SIZE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_SIZE.
        /// </remarks>
        public static readonly PropertyTag MessageSize = new PropertyTag(PropertyId.MessageSize, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_MESSAGE_SUBMISSION_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_SUBMISSION_ID.
        /// </remarks>
        public static readonly PropertyTag MessageSubmissionId = new PropertyTag(PropertyId.MessageSubmissionId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_MESSAGE_TOKEN.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_TOKEN.
        /// </remarks>
        public static readonly PropertyTag MessageToken = new PropertyTag(PropertyId.MessageToken, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_MESSAGE_TO_ME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MESSAGE_TO_ME.
        /// </remarks>
        public static readonly PropertyTag MessageToMe = new PropertyTag(PropertyId.MessageToMe, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_MHS_COMMON_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MHS_COMMON_NAME.
        /// </remarks>
        public static readonly PropertyTag MhsCommonNameA = new PropertyTag(PropertyId.MhsCommonName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_MHS_COMMON_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MHS_COMMON_NAME.
        /// </remarks>
        public static readonly PropertyTag MhsCommonNameW = new PropertyTag(PropertyId.MhsCommonName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_MIDDLE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MIDDLE_NAME.
        /// </remarks>
        public static readonly PropertyTag MiddleNameA = new PropertyTag(PropertyId.MiddleName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_MIDDLE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MIDDLE_NAME.
        /// </remarks>
        public static readonly PropertyTag MiddleNameW = new PropertyTag(PropertyId.MiddleName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_MINI_ICON.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MINI_ICON.
        /// </remarks>
        public static readonly PropertyTag MiniIcon = new PropertyTag(PropertyId.MiniIcon, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag MobileTelephoneNumberA = new PropertyTag(PropertyId.MobileTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag MobileTelephoneNumberW = new PropertyTag(PropertyId.MobileTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_MODIFY_VERSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MODIFY_VERSION.
        /// </remarks>
        public static readonly PropertyTag ModifyVersion = new PropertyTag(PropertyId.ModifyVersion, PropertyType.I8);

        /// <summary>
        /// The MAPI property PR_MSG_STATUS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_MSG_STATUS.
        /// </remarks>
        public static readonly PropertyTag MsgStatus = new PropertyTag(PropertyId.MsgStatus, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_NDR_DIAG_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NDR_DIAG_CODE.
        /// </remarks>
        public static readonly PropertyTag NdrDiagCode = new PropertyTag(PropertyId.NdrDiagCode, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_NDR_REASON_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NDR_REASON_CODE.
        /// </remarks>
        public static readonly PropertyTag NdrReasonCode = new PropertyTag(PropertyId.NdrReasonCode, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_NDR_STATUS_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NDR_STATUS_CODE.
        /// </remarks>
        public static readonly PropertyTag NdrStatusCode = new PropertyTag(PropertyId.NdrStatusCode, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_NEWSGROUP_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NEWSGROUP_NAME.
        /// </remarks>
        public static readonly PropertyTag NewsgroupNameA = new PropertyTag(PropertyId.NewsgroupName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_NEWSGROUP_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NEWSGROUP_NAME.
        /// </remarks>
        public static readonly PropertyTag NewsgroupNameW = new PropertyTag(PropertyId.NewsgroupName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_NICKNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NICKNAME.
        /// </remarks>
        public static readonly PropertyTag NicknameA = new PropertyTag(PropertyId.Nickname, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_NICKNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NICKNAME.
        /// </remarks>
        public static readonly PropertyTag NicknameW = new PropertyTag(PropertyId.Nickname, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_NNTP_XREF.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NNTP_XREF.
        /// </remarks>
        public static readonly PropertyTag NntpXrefA = new PropertyTag(PropertyId.NntpXref, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_NNTP_XREF.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NNTP_XREF.
        /// </remarks>
        public static readonly PropertyTag NntpXrefW = new PropertyTag(PropertyId.NntpXref, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_NON_RECEIPT_NOTIFICATION_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NON_RECEIPT_NOTIFICATION_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag NonReceiptNotificationRequested = new PropertyTag(PropertyId.NonReceiptNotificationRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_NON_RECEIPT_REASON.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NON_RECEIPT_REASON.
        /// </remarks>
        public static readonly PropertyTag NonReceiptReason = new PropertyTag(PropertyId.NonReceiptReason, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_NORMALIZED_SUBJECT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NORMALIZED_SUBJECT.
        /// </remarks>
        public static readonly PropertyTag NormalizedSubjectA = new PropertyTag(PropertyId.NormalizedSubject, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_NORMALIZED_SUBJECT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NORMALIZED_SUBJECT.
        /// </remarks>
        public static readonly PropertyTag NormalizedSubjectW = new PropertyTag(PropertyId.NormalizedSubject, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_NT_SECURITY_DESCRIPTOR.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NT_SECURITY_DESCRIPTOR.
        /// </remarks>
        public static readonly PropertyTag NtSecurityDescriptor = new PropertyTag(PropertyId.NtSecurityDescriptor, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_NULL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_NULL.
        /// </remarks>
        public static readonly PropertyTag Null = new PropertyTag(PropertyId.Null, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_OBJECT_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OBJECT_TYPE.
        /// </remarks>
        public static readonly PropertyTag ObjectType = new PropertyTag(PropertyId.ObjectType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_OBSOLETE_IPMS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OBSOLETE_IPMS.
        /// </remarks>
        public static readonly PropertyTag ObsoletedIpms = new PropertyTag(PropertyId.ObsoletedIpms, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Office2TelephoneNumberA = new PropertyTag(PropertyId.Office2TelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag Office2TelephoneNumberW = new PropertyTag(PropertyId.Office2TelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OFFICE_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OFFICE_LOCATION.
        /// </remarks>
        public static readonly PropertyTag OfficeLocationA = new PropertyTag(PropertyId.OfficeLocation, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OFFICE_LOCATION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OFFICE_LOCATION.
        /// </remarks>
        public static readonly PropertyTag OfficeLocationW = new PropertyTag(PropertyId.OfficeLocation, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag OfficeTelephoneNumberA = new PropertyTag(PropertyId.OfficeTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag OfficeTelephoneNumberW = new PropertyTag(PropertyId.OfficeTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OOF_REPLY_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OOF_REPLY_TYPE.
        /// </remarks>
        public static readonly PropertyTag OofReplyType = new PropertyTag(PropertyId.OofReplyType, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
        /// </remarks>
        public static readonly PropertyTag OrganizationalIdNumberA = new PropertyTag(PropertyId.OrganizationalIdNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
        /// </remarks>
        public static readonly PropertyTag OrganizationalIdNumberW = new PropertyTag(PropertyId.OrganizationalIdNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIG_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIG_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag OrigEntryId = new PropertyTag(PropertyId.OrigEntryId, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorAddrtypeA = new PropertyTag(PropertyId.OriginalAuthorAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorAddrtypeW = new PropertyTag(PropertyId.OriginalAuthorAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorEmailAddressA = new PropertyTag(PropertyId.OriginalAuthorEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorEmailAddressW = new PropertyTag(PropertyId.OriginalAuthorEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorEntryId = new PropertyTag(PropertyId.OriginalAuthorEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorNameA = new PropertyTag(PropertyId.OriginalAuthorName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorNameW = new PropertyTag(PropertyId.OriginalAuthorName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_AUTHOR_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_AUTHOR_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag OriginalAuthorSearchKey = new PropertyTag(PropertyId.OriginalAuthorSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DELIVERY_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DELIVERY_TIME.
        /// </remarks>
        public static readonly PropertyTag OriginalDeliveryTime = new PropertyTag(PropertyId.OriginalDeliveryTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_BCC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_BCC.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayBccA = new PropertyTag(PropertyId.OriginalDisplayBcc, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_BCC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_BCC.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayBccW = new PropertyTag(PropertyId.OriginalDisplayBcc, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_CC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_CC.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayCcA = new PropertyTag(PropertyId.OriginalDisplayCc, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_CC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_CC.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayCcW = new PropertyTag(PropertyId.OriginalDisplayCc, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayNameA = new PropertyTag(PropertyId.OriginalDisplayName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayNameW = new PropertyTag(PropertyId.OriginalDisplayName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_TO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_TO.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayToA = new PropertyTag(PropertyId.OriginalDisplayTo, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_DISPLAY_TO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_DISPLAY_TO.
        /// </remarks>
        public static readonly PropertyTag OriginalDisplayToW = new PropertyTag(PropertyId.OriginalDisplayTo, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_EITS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_EITS.
        /// </remarks>
        public static readonly PropertyTag OriginalEits = new PropertyTag(PropertyId.OriginalEits, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag OriginalEntryId = new PropertyTag(PropertyId.OriginalEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginallyIntendedRecipAddrtypeA = new PropertyTag(PropertyId.OriginallyIntendedRecipAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginallyIntendedRecipAddrtypeW = new PropertyTag(PropertyId.OriginallyIntendedRecipAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginallyIntendedRecipEmailAddressA = new PropertyTag(PropertyId.OriginallyIntendedRecipEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginallyIntendedRecipEmailAddressW = new PropertyTag(PropertyId.OriginallyIntendedRecipEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag OriginallyIntendedRecipEntryId = new PropertyTag(PropertyId.OriginallyIntendedRecipEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIPIENT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINALLY_INTENDED_RECIPIENT_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginallyIntendedRecipientName = new PropertyTag(PropertyId.OriginallyIntendedRecipientName, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag OriginalSearchKey = new PropertyTag(PropertyId.OriginalSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderAddrtypeA = new PropertyTag(PropertyId.OriginalSenderAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderAddrtypeW = new PropertyTag(PropertyId.OriginalSenderAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderEmailAddressA = new PropertyTag(PropertyId.OriginalSenderEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderEmailAddressW = new PropertyTag(PropertyId.OriginalSenderEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderEntryId = new PropertyTag(PropertyId.OriginalSenderEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderNameA = new PropertyTag(PropertyId.OriginalSenderName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderNameW = new PropertyTag(PropertyId.OriginalSenderName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENDER_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENDER_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag OriginalSenderSearchKey = new PropertyTag(PropertyId.OriginalSenderSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENSITIVITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENSITIVITY.
        /// </remarks>
        public static readonly PropertyTag OriginalSensitivity = new PropertyTag(PropertyId.OriginalSensitivity, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingAddrtypeA = new PropertyTag(PropertyId.OriginalSentRepresentingAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingAddrtypeW = new PropertyTag(PropertyId.OriginalSentRepresentingAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingEmailAddressA = new PropertyTag(PropertyId.OriginalSentRepresentingEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingEmailAddressW = new PropertyTag(PropertyId.OriginalSentRepresentingEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingEntryId = new PropertyTag(PropertyId.OriginalSentRepresentingEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingNameA = new PropertyTag(PropertyId.OriginalSentRepresentingName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingNameW = new PropertyTag(PropertyId.OriginalSentRepresentingName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag OriginalSentRepresentingSearchKey = new PropertyTag(PropertyId.OriginalSentRepresentingSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SUBJECT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SUBJECT.
        /// </remarks>
        public static readonly PropertyTag OriginalSubjectA = new PropertyTag(PropertyId.OriginalSubject, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SUBJECT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SUBJECT.
        /// </remarks>
        public static readonly PropertyTag OriginalSubjectW = new PropertyTag(PropertyId.OriginalSubject, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_ORIGINAL_SUBMIT_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINAL_SUBMIT_TIME.
        /// </remarks>
        public static readonly PropertyTag OriginalSubmitTime = new PropertyTag(PropertyId.OriginalSubmitTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_ORIGINATING_MTA_CERTIFICATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINATING_MTA_CERTIFICATE.
        /// </remarks>
        public static readonly PropertyTag OriginatingMtaCertificate = new PropertyTag(PropertyId.OriginatingMtaCertificate, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY.
        /// </remarks>
        public static readonly PropertyTag OriginatorAndDlExpansionHistory = new PropertyTag(PropertyId.OriginatorAndDlExpansionHistory, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINATOR_CERTIFICATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINATOR_CERTIFICATE.
        /// </remarks>
        public static readonly PropertyTag OriginatorCertificate = new PropertyTag(PropertyId.OriginatorCertificate, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag OriginatorDeliveryReportRequested = new PropertyTag(PropertyId.OriginatorDeliveryReportRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag OriginatorNonDeliveryReportRequested = new PropertyTag(PropertyId.OriginatorNonDeliveryReportRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT.
        /// </remarks>
        public static readonly PropertyTag OriginatorRequestedAlternateRecipient = new PropertyTag(PropertyId.OriginatorRequestedAlternateRecipient, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGINATOR_RETURN_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGINATOR_RETURN_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag OriginatorReturnAddress = new PropertyTag(PropertyId.OriginatorReturnAddress, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIGIN_CHECK.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIGIN_CHECK.
        /// </remarks>
        public static readonly PropertyTag OriginCheck = new PropertyTag(PropertyId.OriginCheck, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_ORIG_MESSAGE_CLASS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIG_MESSAGE_CLASS.
        /// </remarks>
        public static readonly PropertyTag OrigMessageClassA = new PropertyTag(PropertyId.OrigMessageClass, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_ORIG_MESSAGE_CLASS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ORIG_MESSAGE_CLASS.
        /// </remarks>
        public static readonly PropertyTag OrigMessageClassW = new PropertyTag(PropertyId.OrigMessageClass, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_CITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_CITY.
        /// </remarks>
        public static readonly PropertyTag OtherAddressCityA = new PropertyTag(PropertyId.OtherAddressCity, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_CITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_CITY.
        /// </remarks>
        public static readonly PropertyTag OtherAddressCityW = new PropertyTag(PropertyId.OtherAddressCity, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag OtherAddressCountryA = new PropertyTag(PropertyId.OtherAddressCountry, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_COUNTRY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_COUNTRY.
        /// </remarks>
        public static readonly PropertyTag OtherAddressCountryW = new PropertyTag(PropertyId.OtherAddressCountry, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag OtherAddressPostalCodeA = new PropertyTag(PropertyId.OtherAddressPostalCode, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag OtherAddressPostalCodeW = new PropertyTag(PropertyId.OtherAddressPostalCode, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
        /// </remarks>
        public static readonly PropertyTag OtherAddressPostOfficeBoxA = new PropertyTag(PropertyId.OtherAddressPostOfficeBox, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
        /// </remarks>
        public static readonly PropertyTag OtherAddressPostOfficeBoxW = new PropertyTag(PropertyId.OtherAddressPostOfficeBox, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
        /// </remarks>
        public static readonly PropertyTag OtherAddressStateOrProvinceA = new PropertyTag(PropertyId.OtherAddressStateOrProvince, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
        /// </remarks>
        public static readonly PropertyTag OtherAddressStateOrProvinceW = new PropertyTag(PropertyId.OtherAddressStateOrProvince, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_STREET.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_STREET.
        /// </remarks>
        public static readonly PropertyTag OtherAddressStreetA = new PropertyTag(PropertyId.OtherAddressStreet, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OTHER_ADDRESS_STREET.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_ADDRESS_STREET.
        /// </remarks>
        public static readonly PropertyTag OtherAddressStreetW = new PropertyTag(PropertyId.OtherAddressStreet, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OTHER_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag OtherTelephoneNumberA = new PropertyTag(PropertyId.OtherTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_OTHER_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OTHER_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag OtherTelephoneNumberW = new PropertyTag(PropertyId.OtherTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_OWNER_APPT_ID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OWNER_APPT_ID.
        /// </remarks>
        public static readonly PropertyTag OwnerApptId = new PropertyTag(PropertyId.OwnerApptId, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_OWN_STORE_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_OWN_STORE_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag OwnStoreEntryId = new PropertyTag(PropertyId.OwnStoreEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PAGER_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PAGER_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag PagerTelephoneNumberA = new PropertyTag(PropertyId.PagerTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PAGER_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PAGER_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag PagerTelephoneNumberW = new PropertyTag(PropertyId.PagerTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PARENT_DISPLAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PARENT_DISPLAY.
        /// </remarks>
        public static readonly PropertyTag ParentDisplayA = new PropertyTag(PropertyId.ParentDisplay, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PARENT_DISPLAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PARENT_DISPLAY.
        /// </remarks>
        public static readonly PropertyTag ParentDisplayW = new PropertyTag(PropertyId.ParentDisplay, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PARENT_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PARENT_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag ParentEntryId = new PropertyTag(PropertyId.ParentEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PARENT_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PARENT_KEY.
        /// </remarks>
        public static readonly PropertyTag ParentKey = new PropertyTag(PropertyId.ParentKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PERSONAL_HOME_PAGE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PERSONAL_HOME_PAGE.
        /// </remarks>
        public static readonly PropertyTag PersonalHomePageA = new PropertyTag(PropertyId.PersonalHomePage, PropertyType.String8);
        /// <summary>
        /// The MAPI property PR_PERSONAL_HOME_PAGE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PERSONAL_HOME_PAGE.
        /// </remarks>
        public static readonly PropertyTag PersonalHomePageW = new PropertyTag(PropertyId.PersonalHomePage, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY.
        /// </remarks>
        public static readonly PropertyTag PhysicalDeliveryBureauFaxDelivery = new PropertyTag(PropertyId.PhysicalDeliveryBureauFaxDelivery, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_PHYSICAL_DELIVERY_MODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PHYSICAL_DELIVERY_MODE.
        /// </remarks>
        public static readonly PropertyTag PhysicalDeliveryMode = new PropertyTag(PropertyId.PhysicalDeliveryMode, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_PHYSICAL_DELIVERY_REPORT_REQUEST.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PHYSICAL_DELIVERY_REPORT_REQUEST.
        /// </remarks>
        public static readonly PropertyTag PhysicalDeliveryReportRequest = new PropertyTag(PropertyId.PhysicalDeliveryReportRequest, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag PhysicalForwardingAddress = new PropertyTag(PropertyId.PhysicalForwardingAddress, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag PhysicalForwardingAddressRequested = new PropertyTag(PropertyId.PhysicalForwardingAddressRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_PHYSICAL_FORWARDING_PROHIBITED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PHYSICAL_FORWARDING_PROHIBITED.
        /// </remarks>
        public static readonly PropertyTag PhysicalForwardingProhibited = new PropertyTag(PropertyId.PhysicalForwardingProhibited, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_PHYSICAL_RENDITION_ATTRIBUTES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PHYSICAL_RENDITION_ATTRIBUTES.
        /// </remarks>
        public static readonly PropertyTag PhysicalRenditionAttributes = new PropertyTag(PropertyId.PhysicalRenditionAttributes, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_POSTAL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POSTAL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag PostalAddressA = new PropertyTag(PropertyId.PostalAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_POSTAL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POSTAL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag PostalAddressW = new PropertyTag(PropertyId.PostalAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag PostalCodeA = new PropertyTag(PropertyId.PostalCode, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_POSTAL_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POSTAL_CODE.
        /// </remarks>
        public static readonly PropertyTag PostalCodeW = new PropertyTag(PropertyId.PostalCode, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_POST_FOLDER_ENTRIES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_FOLDER_ENTRIES.
        /// </remarks>
        public static readonly PropertyTag PostFolderEntries = new PropertyTag(PropertyId.PostFolderEntries, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_POST_FOLDER_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_FOLDER_NAMES.
        /// </remarks>
        public static readonly PropertyTag PostFolderNamesA = new PropertyTag(PropertyId.PostFolderNames, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_POST_FOLDER_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_FOLDER_NAMES.
        /// </remarks>
        public static readonly PropertyTag PostFolderNamesW = new PropertyTag(PropertyId.PostFolderNames, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_POST_OFFICE_BOX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_OFFICE_BOX.
        /// </remarks>
        public static readonly PropertyTag PostOfficeBoxA = new PropertyTag(PropertyId.PostOfficeBox, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_POST_OFFICE_BOX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_OFFICE_BOX.
        /// </remarks>
        public static readonly PropertyTag PostOfficeBoxW = new PropertyTag(PropertyId.PostOfficeBox, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_POST_REPLY_DENIED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_REPLY_DENIED.
        /// </remarks>
        public static readonly PropertyTag PostReplyDenied = new PropertyTag(PropertyId.PostReplyDenied, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_POST_REPLY_FOLDER_ENTRIES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_REPLY_FOLDER_ENTRIES.
        /// </remarks>
        public static readonly PropertyTag PostReplyFolderEntries = new PropertyTag(PropertyId.PostReplyFolderEntries, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_POST_REPLY_FOLDER_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_REPLY_FOLDER_NAMES.
        /// </remarks>
        public static readonly PropertyTag PostReplyFolderNamesA = new PropertyTag(PropertyId.PostReplyFolderNames, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_POST_REPLY_FOLDER_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_POST_REPLY_FOLDER_NAMES.
        /// </remarks>
        public static readonly PropertyTag PostReplyFolderNamesW = new PropertyTag(PropertyId.PostReplyFolderNames, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PREFERRED_BY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PREFERRED_BY_NAME.
        /// </remarks>
        public static readonly PropertyTag PreferredByNameA = new PropertyTag(PropertyId.PreferredByName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PREFERRED_BY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PREFERRED_BY_NAME.
        /// </remarks>
        public static readonly PropertyTag PreferredByNameW = new PropertyTag(PropertyId.PreferredByName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PREPROCESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PREPROCESS.
        /// </remarks>
        public static readonly PropertyTag Preprocess = new PropertyTag(PropertyId.Preprocess, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_PRIMARY_CAPABILITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PRIMARY_CAPABILITY.
        /// </remarks>
        public static readonly PropertyTag PrimaryCapability = new PropertyTag(PropertyId.PrimaryCapability, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PRIMARY_FAX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PRIMARY_FAX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag PrimaryFaxNumberA = new PropertyTag(PropertyId.PrimaryFaxNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PRIMARY_FAX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PRIMARY_FAX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag PrimaryFaxNumberW = new PropertyTag(PropertyId.PrimaryFaxNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag PrimaryTelephoneNumberA = new PropertyTag(PropertyId.PrimaryTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag PrimaryTelephoneNumberW = new PropertyTag(PropertyId.PrimaryTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PRIORITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PRIORITY.
        /// </remarks>
        public static readonly PropertyTag Priority = new PropertyTag(PropertyId.Priority, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_PROFESSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROFESSION.
        /// </remarks>
        public static readonly PropertyTag ProfessionA = new PropertyTag(PropertyId.Profession, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PROFESSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROFESSION.
        /// </remarks>
        public static readonly PropertyTag ProfessionW = new PropertyTag(PropertyId.Profession, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PROFILE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROFILE_NAME.
        /// </remarks>
        public static readonly PropertyTag ProfileNameA = new PropertyTag(PropertyId.ProfileName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PROFILE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROFILE_NAME.
        /// </remarks>
        public static readonly PropertyTag ProfileNameW = new PropertyTag(PropertyId.ProfileName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PROOF_OF_DELIVERY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROOF_OF_DELIVERY.
        /// </remarks>
        public static readonly PropertyTag ProofOfDelivery = new PropertyTag(PropertyId.ProofOfDelivery, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PROOF_OF_DELIVERY_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROOF_OF_DELIVERY_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag ProofOfDeliveryRequested = new PropertyTag(PropertyId.ProofOfDeliveryRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_PROOF_OF_SUBMISSION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROOF_OF_SUBMISSION.
        /// </remarks>
        public static readonly PropertyTag ProofOfSubmission = new PropertyTag(PropertyId.ProofOfSubmission, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PROOF_OF_SUBMISSION_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROOF_OF_SUBMISSION_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag ProofOfSubmissionRequested = new PropertyTag(PropertyId.ProofOfSubmissionRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_PROVIDER_DISPLAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROVIDER_DISPLAY.
        /// </remarks>
        public static readonly PropertyTag ProviderDisplayA = new PropertyTag(PropertyId.ProviderDisplay, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PROVIDER_DISPLAY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROVIDER_DISPLAY.
        /// </remarks>
        public static readonly PropertyTag ProviderDisplayW = new PropertyTag(PropertyId.ProviderDisplay, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PROVIDER_DLL_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROVIDER_DLL_NAME.
        /// </remarks>
        public static readonly PropertyTag ProviderDllNameA = new PropertyTag(PropertyId.ProviderDllName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_PROVIDER_DLL_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROVIDER_DLL_NAME.
        /// </remarks>
        public static readonly PropertyTag ProviderDllNameW = new PropertyTag(PropertyId.ProviderDllName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_PROVIDER_ORDINAL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROVIDER_ORDINAL.
        /// </remarks>
        public static readonly PropertyTag ProviderOrdinal = new PropertyTag(PropertyId.ProviderOrdinal, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_PROVIDER_SUBMIT_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROVIDER_SUBMIT_TIME.
        /// </remarks>
        public static readonly PropertyTag ProviderSubmitTime = new PropertyTag(PropertyId.ProviderSubmitTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_PROVIDER_UID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PROVIDER_UID.
        /// </remarks>
        public static readonly PropertyTag ProviderUid = new PropertyTag(PropertyId.ProviderUid, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_PUID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_PUID.
        /// </remarks>
        public static readonly PropertyTag Puid = new PropertyTag(PropertyId.Puid, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_RADIO_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RADIO_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag RadioTelephoneNumberA = new PropertyTag(PropertyId.RadioTelephoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RADIO_TELEPHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RADIO_TELEPHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag RadioTelephoneNumberW = new PropertyTag(PropertyId.RadioTelephoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingAddrtypeA = new PropertyTag(PropertyId.RcvdRepresentingAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingAddrtypeW = new PropertyTag(PropertyId.RcvdRepresentingAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingEmailAddressA = new PropertyTag(PropertyId.RcvdRepresentingEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingEmailAddressW = new PropertyTag(PropertyId.RcvdRepresentingEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingEntryId = new PropertyTag(PropertyId.RcvdRepresentingEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_NAME.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingNameA = new PropertyTag(PropertyId.RcvdRepresentingName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_NAME.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingNameW = new PropertyTag(PropertyId.RcvdRepresentingName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RCVD_REPRESENTING_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RCVD_REPRESENTING_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag RcvdRepresentingSearchKey = new PropertyTag(PropertyId.RcvdRepresentingSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_READ_RECEIPT_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_READ_RECEIPT_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag ReadReceiptEntryId = new PropertyTag(PropertyId.ReadReceiptEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_READ_RECEIPT_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_READ_RECEIPT_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag ReadReceiptRequested = new PropertyTag(PropertyId.ReadReceiptRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_READ_RECEIPT_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_READ_RECEIPT_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag ReadReceiptSearchKey = new PropertyTag(PropertyId.ReadReceiptSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_RECEIPT_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIPT_TIME.
        /// </remarks>
        public static readonly PropertyTag ReceiptTime = new PropertyTag(PropertyId.ReceiptTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag ReceivedByAddrtypeA = new PropertyTag(PropertyId.ReceivedByAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag ReceivedByAddrtypeW = new PropertyTag(PropertyId.ReceivedByAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag ReceivedByEmailAddressA = new PropertyTag(PropertyId.ReceivedByEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag ReceivedByEmailAddressW = new PropertyTag(PropertyId.ReceivedByEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag ReceivedByEntryId = new PropertyTag(PropertyId.ReceivedByEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_NAME.
        /// </remarks>
        public static readonly PropertyTag ReceivedByNameA = new PropertyTag(PropertyId.ReceivedByName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_NAME.
        /// </remarks>
        public static readonly PropertyTag ReceivedByNameW = new PropertyTag(PropertyId.ReceivedByName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RECEIVED_BY_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVED_BY_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag ReceivedBySearchKey = new PropertyTag(PropertyId.ReceivedBySearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_RECEIVE_FOLDER_SETTINGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECEIVE_FOLDER_SETTINGS.
        /// </remarks>
        public static readonly PropertyTag ReceiveFolderSettings = new PropertyTag(PropertyId.ReceiveFolderSettings, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_CERTIFICATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_CERTIFICATE.
        /// </remarks>
        public static readonly PropertyTag RecipientCertificate = new PropertyTag(PropertyId.RecipientCertificate, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag RecipientDisplayNameA = new PropertyTag(PropertyId.RecipientDisplayName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag RecipientDisplayNameW = new PropertyTag(PropertyId.RecipientDisplayName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
        /// </remarks>
        public static readonly PropertyTag RecipientNumberForAdviceA = new PropertyTag(PropertyId.RecipientNumberForAdvice, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
        /// </remarks>
        public static readonly PropertyTag RecipientNumberForAdviceW = new PropertyTag(PropertyId.RecipientNumberForAdvice, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_REASSIGNMENT_PROHIBITED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_REASSIGNMENT_PROHIBITED.
        /// </remarks>
        public static readonly PropertyTag RecipientReassignmentProhibited = new PropertyTag(PropertyId.RecipientReassignmentProhibited, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_STATUS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_STATUS.
        /// </remarks>
        public static readonly PropertyTag RecipientStatus = new PropertyTag(PropertyId.RecipientStatus, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RECIPIENT_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RECIPIENT_TYPE.
        /// </remarks>
        public static readonly PropertyTag RecipientType = new PropertyTag(PropertyId.RecipientType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_REDIRECTION_HISTORY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REDIRECTION_HISTORY.
        /// </remarks>
        public static readonly PropertyTag RedirectionHistory = new PropertyTag(PropertyId.RedirectionHistory, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REFERRED_BY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REFERRED_BY_NAME.
        /// </remarks>
        public static readonly PropertyTag ReferredByNameA = new PropertyTag(PropertyId.ReferredByName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_REFERRED_BY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REFERRED_BY_NAME.
        /// </remarks>
        public static readonly PropertyTag ReferredByNameW = new PropertyTag(PropertyId.ReferredByName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_REGISTERED_MAIL_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REGISTERED_MAIL_TYPE.
        /// </remarks>
        public static readonly PropertyTag RegisteredMailType = new PropertyTag(PropertyId.RegisteredMailType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RELATED_IPMS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RELATED_IPMS.
        /// </remarks>
        public static readonly PropertyTag RelatedIpms = new PropertyTag(PropertyId.RelatedIpms, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REMOTE_PROGRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REMOTE_PROGRESS.
        /// </remarks>
        public static readonly PropertyTag RemoteProgress = new PropertyTag(PropertyId.RemoteProgress, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_REMOTE_PROGRESS_TEXT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REMOTE_PROGRESS_TEXT.
        /// </remarks>
        public static readonly PropertyTag RemoteProgressTextA = new PropertyTag(PropertyId.RemoteProgressText, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_REMOTE_PROGRESS_TEXT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REMOTE_PROGRESS_TEXT.
        /// </remarks>
        public static readonly PropertyTag RemoteProgressTextW = new PropertyTag(PropertyId.RemoteProgressText, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_REMOTE_VALIDATE_OK.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REMOTE_VALIDATE_OK.
        /// </remarks>
        public static readonly PropertyTag RemoteValidateOk = new PropertyTag(PropertyId.RemoteValidateOk, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_RENDERING_POSITION.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RENDERING_POSITION.
        /// </remarks>
        public static readonly PropertyTag RenderingPosition = new PropertyTag(PropertyId.RenderingPosition, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_REPLY_RECIPIENT_ENTRIES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPLY_RECIPIENT_ENTRIES.
        /// </remarks>
        public static readonly PropertyTag ReplyRecipientEntries = new PropertyTag(PropertyId.ReplyRecipientEntries, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REPLY_RECIPIENT_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPLY_RECIPIENT_NAMES.
        /// </remarks>
        public static readonly PropertyTag ReplyRecipientNamesA = new PropertyTag(PropertyId.ReplyRecipientNames, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_REPLY_RECIPIENT_NAMES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPLY_RECIPIENT_NAMES.
        /// </remarks>
        public static readonly PropertyTag ReplyRecipientNamesW = new PropertyTag(PropertyId.ReplyRecipientNames, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_REPLY_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPLY_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag ReplyRequested = new PropertyTag(PropertyId.ReplyRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_REPLY_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPLY_TIME.
        /// </remarks>
        public static readonly PropertyTag ReplyTime = new PropertyTag(PropertyId.ReplyTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_REPORT_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag ReportEntryId = new PropertyTag(PropertyId.ReportEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REPORTING_DL_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORTING_DL_NAME.
        /// </remarks>
        public static readonly PropertyTag ReportingDlName = new PropertyTag(PropertyId.ReportingDlName, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REPORTING_MTA_CERTIFICATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORTING_MTA_CERTIFICATE.
        /// </remarks>
        public static readonly PropertyTag ReportingMtaCertificate = new PropertyTag(PropertyId.ReportingMtaCertificate, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REPORT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_NAME.
        /// </remarks>
        public static readonly PropertyTag ReportNameA = new PropertyTag(PropertyId.ReportName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_REPORT_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_NAME.
        /// </remarks>
        public static readonly PropertyTag ReportNameW = new PropertyTag(PropertyId.ReportName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_REPORT_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag ReportSearchKey = new PropertyTag(PropertyId.ReportSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REPORT_TAG.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_TAG.
        /// </remarks>
        public static readonly PropertyTag ReportTag = new PropertyTag(PropertyId.ReportTag, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_REPORT_TEXT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_TEXT.
        /// </remarks>
        public static readonly PropertyTag ReportTextA = new PropertyTag(PropertyId.ReportText, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_REPORT_TEXT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_TEXT.
        /// </remarks>
        public static readonly PropertyTag ReportTextW = new PropertyTag(PropertyId.ReportText, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_REPORT_TIME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REPORT_TIME.
        /// </remarks>
        public static readonly PropertyTag ReportTime = new PropertyTag(PropertyId.ReportTime, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_REQUESTED_DELIVERY_METHOD.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_REQUESTED_DELIVERY_METHOD.
        /// </remarks>
        public static readonly PropertyTag RequestedDeliveryMethod = new PropertyTag(PropertyId.RequestedDeliveryMethod, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RESOURCE_FLAGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RESOURCE_FLAGS.
        /// </remarks>
        public static readonly PropertyTag ResourceFlags = new PropertyTag(PropertyId.ResourceFlags, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RESOURCE_METHODS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RESOURCE_METHODS.
        /// </remarks>
        public static readonly PropertyTag ResourceMethods = new PropertyTag(PropertyId.ResourceMethods, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RESOURCE_PATH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RESOURCE_PATH.
        /// </remarks>
        public static readonly PropertyTag ResourcePathA = new PropertyTag(PropertyId.ResourcePath, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RESOURCE_PATH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RESOURCE_PATH.
        /// </remarks>
        public static readonly PropertyTag ResourcePathW = new PropertyTag(PropertyId.ResourcePath, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RESOURCE_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RESOURCE_TYPE.
        /// </remarks>
        public static readonly PropertyTag ResourceType = new PropertyTag(PropertyId.ResourceType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RESPONSE_REQUESTED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RESPONSE_REQUESTED.
        /// </remarks>
        public static readonly PropertyTag ResponseRequested = new PropertyTag(PropertyId.ResponseRequested, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_RESPONSIBILITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RESPONSIBILITY.
        /// </remarks>
        public static readonly PropertyTag Responsibility = new PropertyTag(PropertyId.Responsibility, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_RETURNED_IPM.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RETURNED_IPM.
        /// </remarks>
        public static readonly PropertyTag ReturnedIpm = new PropertyTag(PropertyId.ReturnedIpm, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_ROWID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ROWID.
        /// </remarks>
        public static readonly PropertyTag Rowid = new PropertyTag(PropertyId.Rowid, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_ROW_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_ROW_TYPE.
        /// </remarks>
        public static readonly PropertyTag RowType = new PropertyTag(PropertyId.RowType, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RTF_COMPRESSED.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_COMPRESSED.
        /// </remarks>
        public static readonly PropertyTag RtfCompressed = new PropertyTag(PropertyId.RtfCompressed, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_RTF_IN_SYNC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_IN_SYNC.
        /// </remarks>
        public static readonly PropertyTag RtfInSync = new PropertyTag(PropertyId.RtfInSync, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_RTF_SYNC_BODY_COUNT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_SYNC_BODY_COUNT.
        /// </remarks>
        public static readonly PropertyTag RtfSyncBodyCount = new PropertyTag(PropertyId.RtfSyncBodyCount, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RTF_SYNC_BODY_CRC.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_SYNC_BODY_CRC.
        /// </remarks>
        public static readonly PropertyTag RtfSyncBodyCrc = new PropertyTag(PropertyId.RtfSyncBodyCrc, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RTF_SYNC_BODY_TAG.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_SYNC_BODY_TAG.
        /// </remarks>
        public static readonly PropertyTag RtfSyncBodyTagA = new PropertyTag(PropertyId.RtfSyncBodyTag, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_RTF_SYNC_BODY_TAG.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_SYNC_BODY_TAG.
        /// </remarks>
        public static readonly PropertyTag RtfSyncBodyTagW = new PropertyTag(PropertyId.RtfSyncBodyTag, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_RTF_SYNC_PREFIX_COUNT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_SYNC_PREFIX_COUNT.
        /// </remarks>
        public static readonly PropertyTag RtfSyncPrefixCount = new PropertyTag(PropertyId.RtfSyncPrefixCount, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_RTF_SYNC_TRAILING_COUNT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_RTF_SYNC_TRAILING_COUNT.
        /// </remarks>
        public static readonly PropertyTag RtfSyncTrailingCount = new PropertyTag(PropertyId.RtfSyncTrailingCount, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_SEARCH.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SEARCH.
        /// </remarks>
        public static readonly PropertyTag Search = new PropertyTag(PropertyId.Search, PropertyType.Object);

        /// <summary>
        /// The MAPI property PR_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag SearchKey = new PropertyTag(PropertyId.SearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SECURITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SECURITY.
        /// </remarks>
        public static readonly PropertyTag Security = new PropertyTag(PropertyId.Security, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_SELECTABLE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SELECTABLE.
        /// </remarks>
        public static readonly PropertyTag Selectable = new PropertyTag(PropertyId.Selectable, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_SENDER_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag SenderAddrtypeA = new PropertyTag(PropertyId.SenderAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SENDER_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag SenderAddrtypeW = new PropertyTag(PropertyId.SenderAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SENDER_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag SenderEmailAddressA = new PropertyTag(PropertyId.SenderEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SENDER_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag SenderEmailAddressW = new PropertyTag(PropertyId.SenderEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SENDER_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag SenderEntryId = new PropertyTag(PropertyId.SenderEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SENDER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_NAME.
        /// </remarks>
        public static readonly PropertyTag SenderNameA = new PropertyTag(PropertyId.SenderName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SENDER_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_NAME.
        /// </remarks>
        public static readonly PropertyTag SenderNameW = new PropertyTag(PropertyId.SenderName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SENDER_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENDER_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag SenderSearchKey = new PropertyTag(PropertyId.SenderSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SEND_INTERNET_ENCODING.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SEND_INTERNET_ENCODING.
        /// </remarks>
        public static readonly PropertyTag SendInternetEncoding = new PropertyTag(PropertyId.SendInternetEncoding, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_SEND_RECALL_REPORT
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SEND_RECALL_REPORT.
        /// </remarks>
        public static readonly PropertyTag SendRecallReport = new PropertyTag(PropertyId.SendRecallReport, PropertyType.Unspecified);

        /// <summary>
        /// The MAPI property PR_SEND_RICH_INFO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SEND_RICH_INFO.
        /// </remarks>
        public static readonly PropertyTag SendRichInfo = new PropertyTag(PropertyId.SendRichInfo, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_SENSITIVITY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENSITIVITY.
        /// </remarks>
        public static readonly PropertyTag Sensitivity = new PropertyTag(PropertyId.Sensitivity, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_SENTMAIL_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENTMAIL_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag SentmailEntryId = new PropertyTag(PropertyId.SentmailEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingAddrtypeA = new PropertyTag(PropertyId.SentRepresentingAddrtype, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingAddrtypeW = new PropertyTag(PropertyId.SentRepresentingAddrtype, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingEmailAddressA = new PropertyTag(PropertyId.SentRepresentingEmailAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingEmailAddressW = new PropertyTag(PropertyId.SentRepresentingEmailAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingEntryId = new PropertyTag(PropertyId.SentRepresentingEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_NAME.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingNameA = new PropertyTag(PropertyId.SentRepresentingName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_NAME.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingNameW = new PropertyTag(PropertyId.SentRepresentingName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SENT_REPRESENTING_SEARCH_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SENT_REPRESENTING_SEARCH_KEY.
        /// </remarks>
        public static readonly PropertyTag SentRepresentingSearchKey = new PropertyTag(PropertyId.SentRepresentingSearchKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SERVICE_DELETE_FILES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_DELETE_FILES.
        /// </remarks>
        public static readonly PropertyTag ServiceDeleteFilesA = new PropertyTag(PropertyId.ServiceDeleteFiles, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SERVICE_DELETE_FILES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_DELETE_FILES.
        /// </remarks>
        public static readonly PropertyTag ServiceDeleteFilesW = new PropertyTag(PropertyId.ServiceDeleteFiles, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SERVICE_DLL_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_DLL_NAME.
        /// </remarks>
        public static readonly PropertyTag ServiceDllNameA = new PropertyTag(PropertyId.ServiceDllName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SERVICE_DLL_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_DLL_NAME.
        /// </remarks>
        public static readonly PropertyTag ServiceDllNameW = new PropertyTag(PropertyId.ServiceDllName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SERVICE_ENTRY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_ENTRY_NAME.
        /// </remarks>
        public static readonly PropertyTag ServiceEntryName = new PropertyTag(PropertyId.ServiceEntryName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SERVICE_EXTRA_UIDS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_EXTRA_UIDS.
        /// </remarks>
        public static readonly PropertyTag ServiceExtraUids = new PropertyTag(PropertyId.ServiceExtraUids, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SERVICE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_NAME.
        /// </remarks>
        public static readonly PropertyTag ServiceNameA = new PropertyTag(PropertyId.ServiceName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SERVICE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_NAME.
        /// </remarks>
        public static readonly PropertyTag ServiceNameW = new PropertyTag(PropertyId.ServiceName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SERVICES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICES.
        /// </remarks>
        public static readonly PropertyTag Services = new PropertyTag(PropertyId.Services, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SERVICE_SUPPORT_FILES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_SUPPORT_FILES.
        /// </remarks>
        public static readonly PropertyTag ServiceSupportFilesA = new PropertyTag(PropertyId.ServiceSupportFiles, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SERVICE_SUPPORT_FILES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_SUPPORT_FILES.
        /// </remarks>
        public static readonly PropertyTag ServiceSupportFilesW = new PropertyTag(PropertyId.ServiceSupportFiles, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SERVICE_UID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SERVICE_UID.
        /// </remarks>
        public static readonly PropertyTag ServiceUid = new PropertyTag(PropertyId.ServiceUid, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SEVEN_BIT_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SEVEN_BIT_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag SevenBitDisplayName = new PropertyTag(PropertyId.SevenBitDisplayName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SMTP_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SMTP_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag SmtpAddressA = new PropertyTag(PropertyId.SmtpAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SMTP_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SMTP_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag SmtpAddressW = new PropertyTag(PropertyId.SmtpAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SPOOLER_STATUS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SPOOLER_STATUS.
        /// </remarks>
        public static readonly PropertyTag SpoolerStatus = new PropertyTag(PropertyId.SpoolerStatus, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_SPOUSE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SPOUSE_NAME.
        /// </remarks>
        public static readonly PropertyTag SpouseNameA = new PropertyTag(PropertyId.SpouseName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SPOUSE_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SPOUSE_NAME.
        /// </remarks>
        public static readonly PropertyTag SpouseNameW = new PropertyTag(PropertyId.SpouseName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_START_DATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_START_DATE.
        /// </remarks>
        public static readonly PropertyTag StartDate = new PropertyTag(PropertyId.StartDate, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_STATE_OR_PROVINCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STATE_OR_PROVINCE.
        /// </remarks>
        public static readonly PropertyTag StateOrProvinceA = new PropertyTag(PropertyId.StateOrProvince, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_STATE_OR_PROVINCE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STATE_OR_PROVINCE.
        /// </remarks>
        public static readonly PropertyTag StateOrProvinceW = new PropertyTag(PropertyId.StateOrProvince, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_STATUS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STATUS.
        /// </remarks>
        public static readonly PropertyTag Status = new PropertyTag(PropertyId.Status, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_STATUS_CODE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STATUS_CODE.
        /// </remarks>
        public static readonly PropertyTag StatusCode = new PropertyTag(PropertyId.StatusCode, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_STATUS_STRING.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STATUS_STRING.
        /// </remarks>
        public static readonly PropertyTag StatusStringA = new PropertyTag(PropertyId.StatusString, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_STATUS_STRING.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STATUS_STRING.
        /// </remarks>
        public static readonly PropertyTag StatusStringW = new PropertyTag(PropertyId.StatusString, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_STORE_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STORE_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag StoreEntryId = new PropertyTag(PropertyId.StoreEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_STORE_PROVIDERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STORE_PROVIDERS.
        /// </remarks>
        public static readonly PropertyTag StoreProviders = new PropertyTag(PropertyId.StoreProviders, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_STORE_RECORD_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STORE_RECORD_KEY.
        /// </remarks>
        public static readonly PropertyTag StoreRecordKey = new PropertyTag(PropertyId.StoreRecordKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_STORE_STATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STORE_STATE.
        /// </remarks>
        public static readonly PropertyTag StoreState = new PropertyTag(PropertyId.StoreState, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_STORE_SUPPORT_MASK.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STORE_SUPPORT_MASK.
        /// </remarks>
        public static readonly PropertyTag StoreSupportMask = new PropertyTag(PropertyId.StoreSupportMask, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_STREET_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STREET_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag StreetAddressA = new PropertyTag(PropertyId.StreetAddress, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_STREET_ADDRESS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_STREET_ADDRESS.
        /// </remarks>
        public static readonly PropertyTag StreetAddressW = new PropertyTag(PropertyId.StreetAddress, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SUBFOLDERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUBFOLDERS.
        /// </remarks>
        public static readonly PropertyTag Subfolders = new PropertyTag(PropertyId.Subfolders, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_SUBJECT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUBJECT.
        /// </remarks>
        public static readonly PropertyTag SubjectA = new PropertyTag(PropertyId.Subject, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SUBJECT.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUBJECT.
        /// </remarks>
        public static readonly PropertyTag SubjectW = new PropertyTag(PropertyId.Subject, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SUBJECT_IPM.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUBJECT_IPM.
        /// </remarks>
        public static readonly PropertyTag SubjectIpm = new PropertyTag(PropertyId.SubjectIpm, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_SUBJECT_PREFIX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUBJECT_PREFIX.
        /// </remarks>
        public static readonly PropertyTag SubjectPrefixA = new PropertyTag(PropertyId.SubjectPrefix, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SUBJECT_PREFIX.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUBJECT_PREFIX.
        /// </remarks>
        public static readonly PropertyTag SubjectPrefixW = new PropertyTag(PropertyId.SubjectPrefix, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SUBMIT_FLAGS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUBMIT_FLAGS.
        /// </remarks>
        public static readonly PropertyTag SubmitFlags = new PropertyTag(PropertyId.SubmitFlags, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_SUPERSEDES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUPERSEDES.
        /// </remarks>
        public static readonly PropertyTag SupersedesA = new PropertyTag(PropertyId.Supersedes, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SUPERSEDES.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUPERSEDES.
        /// </remarks>
        public static readonly PropertyTag SupersedesW = new PropertyTag(PropertyId.Supersedes, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SUPPLEMENTARY_INFO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUPPLEMENTARY_INFO.
        /// </remarks>
        public static readonly PropertyTag SupplementaryInfoA = new PropertyTag(PropertyId.SupplementaryInfo, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SUPPLEMENTARY_INFO.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SUPPLEMENTARY_INFO.
        /// </remarks>
        public static readonly PropertyTag SupplementaryInfoW = new PropertyTag(PropertyId.SupplementaryInfo, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_SURNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SURNAME.
        /// </remarks>
        public static readonly PropertyTag SurnameA = new PropertyTag(PropertyId.Surname, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_SURNAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_SURNAME.
        /// </remarks>
        public static readonly PropertyTag SurnameW = new PropertyTag(PropertyId.Surname, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_TELEX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TELEX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag TelexNumberA = new PropertyTag(PropertyId.TelexNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_TELEX_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TELEX_NUMBER.
        /// </remarks>
        public static readonly PropertyTag TelexNumberW = new PropertyTag(PropertyId.TelexNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_TEMPLATEID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TEMPLATEID.
        /// </remarks>
        public static readonly PropertyTag Templateid = new PropertyTag(PropertyId.Templateid, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_TITLE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TITLE.
        /// </remarks>
        public static readonly PropertyTag TitleA = new PropertyTag(PropertyId.Title, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_TITLE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TITLE.
        /// </remarks>
        public static readonly PropertyTag TitleW = new PropertyTag(PropertyId.Title, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_TNEF_CORRELATION_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TNEF_CORRELATION_KEY.
        /// </remarks>
        public static readonly PropertyTag TnefCorrelationKey = new PropertyTag(PropertyId.TnefCorrelationKey, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag TransmitableDisplayNameA = new PropertyTag(PropertyId.TransmitableDisplayName, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
        /// </remarks>
        public static readonly PropertyTag TransmitableDisplayNameW = new PropertyTag(PropertyId.TransmitableDisplayName, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_TRANSPORT_KEY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TRANSPORT_KEY.
        /// </remarks>
        public static readonly PropertyTag TransportKey = new PropertyTag(PropertyId.TransportKey, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
        /// </remarks>
        public static readonly PropertyTag TransportMessageHeadersA = new PropertyTag(PropertyId.TransportMessageHeaders, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
        /// </remarks>
        public static readonly PropertyTag TransportMessageHeadersW = new PropertyTag(PropertyId.TransportMessageHeaders, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_TRANSPORT_PROVIDERS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TRANSPORT_PROVIDERS.
        /// </remarks>
        public static readonly PropertyTag TransportProviders = new PropertyTag(PropertyId.TransportProviders, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_TRANSPORT_STATUS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TRANSPORT_STATUS.
        /// </remarks>
        public static readonly PropertyTag TransportStatus = new PropertyTag(PropertyId.TransportStatus, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_TTYDD_PHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TTYDD_PHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag TtytddPhoneNumberA = new PropertyTag(PropertyId.TtytddPhoneNumber, PropertyType.String8);

        /// <summary>
        /// The MAPI property PR_TTYDD_PHONE_NUMBER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TTYDD_PHONE_NUMBER.
        /// </remarks>
        public static readonly PropertyTag TtytddPhoneNumberW = new PropertyTag(PropertyId.TtytddPhoneNumber, PropertyType.Unicode);

        /// <summary>
        /// The MAPI property PR_TYPE_OF_MTS_USER.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_TYPE_OF_MTS_USER.
        /// </remarks>
        public static readonly PropertyTag TypeOfMtsUser = new PropertyTag(PropertyId.TypeOfMtsUser, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_USER_CERTIFICATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_USER_CERTIFICATE.
        /// </remarks>
        public static readonly PropertyTag UserCertificate = new PropertyTag(PropertyId.UserCertificate, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_USER_X509_CERTIFICATE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_USER_X509_CERTIFICATE.
        /// </remarks>
        public static readonly PropertyTag UserX509Certificate = new PropertyTag(PropertyId.UserX509Certificate, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_VALID_FOLDER_MASK.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_VALID_FOLDER_MASK.
        /// </remarks>
        public static readonly PropertyTag ValidFolderMask = new PropertyTag(PropertyId.ValidFolderMask, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_VIEWS_ENTRYID.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_VIEWS_ENTRYID.
        /// </remarks>
        public static readonly PropertyTag ViewsEntryId = new PropertyTag(PropertyId.ViewsEntryId, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_WEDDING_ANNIVERSARY.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_WEDDING_ANNIVERSARY.
        /// </remarks>
        public static readonly PropertyTag WeddingAnniversary = new PropertyTag(PropertyId.WeddingAnniversary, PropertyType.SysTime);

        /// <summary>
        /// The MAPI property PR_X400_CONTENT_TYPE.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_X400_CONTENT_TYPE.
        /// </remarks>
        public static readonly PropertyTag X400ContentType = new PropertyTag(PropertyId.X400ContentType, PropertyType.Binary);

        /// <summary>
        /// The MAPI property PR_X400_DEFERRED_DELIVERY_CANCEL.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_X400_DEFERRED_DELIVERY_CANCEL.
        /// </remarks>
        public static readonly PropertyTag X400DeferredDeliveryCancel = new PropertyTag(PropertyId.X400DeferredDeliveryCancel, PropertyType.Boolean);

        /// <summary>
        /// The MAPI property PR_XPOS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_XPOS.
        /// </remarks>
        public static readonly PropertyTag Xpos = new PropertyTag(PropertyId.Xpos, PropertyType.Long);

        /// <summary>
        /// The MAPI property PR_YPOS.
        /// </summary>
        /// <remarks>
        /// The MAPI property PR_YPOS.
        /// </remarks>
        public static readonly PropertyTag Ypos = new PropertyTag(PropertyId.Ypos, PropertyType.Long);

        private const PropertyId NamedMin = unchecked((PropertyId)0x8000);
        private const PropertyId NamedMax = unchecked((PropertyId)0xFFFE);
        private const short MultiValuedFlag = (short)PropertyType.MultiValued;
        private readonly PropertyType type;
        private readonly PropertyId id;

        /// <summary>
        /// Get the property identifier.
        /// </summary>
        /// <remarks>
        /// Gets the property identifier.
        /// </remarks>
        /// <value>The identifier.</value>
        public PropertyId Id
        {
            get { return id; }
        }

        /// <summary>
        /// Get a value indicating whether or not the property contains multiple values.
        /// </summary>
        /// <remarks>
        /// Gets a value indicating whether or not the property contains multiple values.
        /// </remarks>
        /// <value><c>true</c> if the property contains multiple values; otherwise, <c>false</c>.</value>
        public bool IsMultiValued
        {
            get { return ((short)type & MultiValuedFlag) != 0; }
        }

        /// <summary>
        /// Get a value indicating whether or not the property has a special name.
        /// </summary>
        /// <remarks>
        /// Gets a value indicating whether or not the property has a special name.
        /// </remarks>
        /// <value><c>true</c> if the property has a special name; otherwise, <c>false</c>.</value>
        public bool IsNamed
        {
            get { return (int)id >= (int)NamedMin && (int)id <= (int)NamedMax; }
        }

        /// <summary>
        /// Get a value indicating whether the property value type is valid.
        /// </summary>
        /// <remarks>
        /// Gets a value indicating whether the property value type is valid.
        /// </remarks>
        /// <value><c>true</c> if the property value type is valid; otherwise, <c>false</c>.</value>
        public bool IsTypeValid
        {
            get
            {
                switch (ValueTnefType)
                {
                    case PropertyType.Unspecified:
                    case PropertyType.Null:
                    case PropertyType.I2:
                    case PropertyType.Long:
                    case PropertyType.R4:
                    case PropertyType.Double:
                    case PropertyType.Currency:
                    case PropertyType.AppTime:
                    case PropertyType.Error:
                    case PropertyType.Boolean:
                    case PropertyType.Object:
                    case PropertyType.I8:
                    case PropertyType.String8:
                    case PropertyType.Unicode:
                    case PropertyType.SysTime:
                    case PropertyType.ClassId:
                    case PropertyType.Binary:
                        return true;
                    default:
                        return false;
                }
            }
        }

        /// <summary>
        /// Get the property's value type (including the multi-valued bit).
        /// </summary>
        /// <remarks>
        /// Gets the property's value type (including the multi-valued bit).
        /// </remarks>
        /// <value>The property's value type.</value>
        public PropertyType TnefType
        {
            get { return type; }
        }

        /// <summary>
        /// Get the type of the value that the property contains.
        /// </summary>
        /// <remarks>
        /// Gets the type of the value that the property contains.
        /// </remarks>
        /// <value>The type of the value.</value>
        public PropertyType ValueTnefType
        {
            get { return (PropertyType)((short)type & ~MultiValuedFlag); }
        }

        /// <summary>
        /// Initialize a new instance of the <see cref="PropertyTag"/> struct.
        /// </summary>
        /// <remarks>
        /// Creates a new <see cref="PropertyTag"/> based on a 32-bit integer tag as read from
        /// a TNEF stream.
        /// </remarks>
        /// <param name="tag">The property tag.</param>
        public PropertyTag(int tag)
        {
            type = (PropertyType)(tag >> 16 & 0xFFFF);
            id = (PropertyId)(tag & 0xFFFF);
        }

        PropertyTag(PropertyId id, PropertyType type, bool multiValue)
        {
            this.type = (PropertyType)((ushort)type | (multiValue ? MultiValuedFlag : 0));
            this.id = id;
        }

        /// <summary>
        /// Initialize a new instance of the <see cref="PropertyTag"/> struct.
        /// </summary>
        /// <remarks>
        /// Creates a new <see cref="PropertyTag"/> based on a <see cref="PropertyId"/>
        /// and <see cref="PropertyType"/>.
        /// </remarks>
        /// <param name="id">The property identifier.</param>
        /// <param name="type">The property type.</param>
        public PropertyTag(PropertyId id, PropertyType type)
        {
            this.type = type;
            this.id = id;
        }

        /// <summary>
        /// Cast an integer tag value into a TNEF property tag.
        /// </summary>
        /// <remarks>
        /// Casts an integer tag value into a TNEF property tag.
        /// </remarks>
        /// <returns>A <see cref="PropertyTag"/> that represents the integer tag value.</returns>
        /// <param name="tag">The integer tag value.</param>
        public static implicit operator PropertyTag(int tag)
        {
            return new PropertyTag(tag);
        }

        /// <summary>
        /// Cast a TNEF property tag into a 32-bit integer value.
        /// </summary>
        /// <remarks>
        /// Casts a TNEF property tag into a 32-bit integer value.
        /// </remarks>
        /// <returns>A 32-bit integer value representing the TNEF property tag.</returns>
        /// <param name="tag">The TNEF property tag.</param>
        public static implicit operator int(PropertyTag tag)
        {
            return (ushort)tag.TnefType << 16 | (ushort)tag.Id;
        }

        /// <summary>
        /// Serves as a hash function for a <see cref="PropertyTag"/> object.
        /// </summary>
        /// <remarks>
        /// Serves as a hash function for a <see cref="PropertyTag"/> object.
        /// </remarks>
        /// <returns>A hash code for this instance that is suitable for use in hashing algorithms
        /// and data structures such as a hash table.</returns>
        public override int GetHashCode()
        {
            return ((int)this).GetHashCode();
        }

        /// <summary>
        /// Determine whether the specified <see cref="object"/> is equal to the current <see cref="PropertyTag"/>.
        /// </summary>
        /// <remarks>
        /// Determines whether the specified <see cref="object"/> is equal to the current <see cref="PropertyTag"/>.
        /// </remarks>
        /// <param name="obj">The <see cref="object"/> to compare with the current <see cref="PropertyTag"/>.</param>
        /// <returns><c>true</c> if the specified <see cref="object"/> is equal to the current
        /// <see cref="PropertyTag"/>; otherwise, <c>false</c>.</returns>
        public override bool Equals(object obj)
        {
            return obj is PropertyTag tag
                && tag.Id == Id
                && tag.TnefType == TnefType;
        }

        /// <summary>
        /// Return a <see cref="string"/> that represents the current <see cref="PropertyTag"/>.
        /// </summary>
        /// <remarks>
        /// Returns a <see cref="string"/> that represents the current <see cref="PropertyTag"/>.
        /// </remarks>
        /// <returns>A <see cref="string"/> that represents the current <see cref="PropertyTag"/>.</returns>
        public override string ToString()
        {
            return $"{Id} ({ValueTnefType})";
        }

        /// <summary>
        /// Return a new <see cref="PropertyTag"/> where the type has been changed to <see cref="PropertyType.Unicode"/>.
        /// </summary>
        /// <remarks>
        /// Returns a new <see cref="PropertyTag"/> where the type has been changed to <see cref="PropertyType.Unicode"/>.
        /// </remarks>
        /// <returns>The unicode equivalent of the property tag.</returns>
        public PropertyTag ToUnicode()
        {
            var unicode = (PropertyType)((short)type & MultiValuedFlag | (short)PropertyType.Unicode);

            return new PropertyTag(id, unicode);
        }
    }
}
