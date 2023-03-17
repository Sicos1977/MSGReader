//
// PropertyTag.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2022 .NET Foundation and Contributors
//
// Refactoring to the code done by Kees van Spelde so that it works in this project
// Copyright (c) 2023 Kees van Spelde <sicos2002@hotmail.com>
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

using MsgReader.Tnef.Enums;
// ReSharper disable All

namespace MsgReader.Tnef;

/// <summary>
///     A TNEF property tag.
/// </summary>
/// <remarks>
///     A TNEF property tag.
/// </remarks>
public readonly struct PropertyTag
{
    #region Statics
    /// <summary>
    ///     The MAPI property PR_AB_DEFAULT_DIR.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AB_DEFAULT_DIR.
    /// </remarks>
    public static readonly PropertyTag AbDefaultDir = new(PropertyId.AbDefaultDir, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_AB_DEFAULT_PAB.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AB_DEFAULT_PAD.
    /// </remarks>
    public static readonly PropertyTag AbDefaultPab = new(PropertyId.AbDefaultPab, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_AB_PROVIDER_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AB_PROVIDER_ID.
    /// </remarks>
    public static readonly PropertyTag AbProviderId = new(PropertyId.AbProviderId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_AB_PROVIDERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AB_PROVIDERS.
    /// </remarks>
    public static readonly PropertyTag AbProviders = new(PropertyId.AbProviders, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_AB_SEARCH_PATH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AB_SEARCH_PATH.
    /// </remarks>
    public static readonly PropertyTag AbSearchPath = new(PropertyId.AbSearchPath, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_AB_SEARCH_PATH_UPDATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AB_SEARCH_PATH_UPDATE.
    /// </remarks>
    public static readonly PropertyTag AbSearchPathUpdate = new(PropertyId.AbSearchPathUpdate, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ACCESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ACCESS.
    /// </remarks>
    public static readonly PropertyTag Access = new(PropertyId.Access, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ACCESS_LEVEL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ACCESS_LEVEL.
    /// </remarks>
    public static readonly PropertyTag AccessLevel = new(PropertyId.AccessLevel, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ACCOUNT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ACCOUNT.
    /// </remarks>
    public static readonly PropertyTag AccountA = new(PropertyId.Account, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ACCOUNT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ACCOUNT.
    /// </remarks>
    public static readonly PropertyTag AccountW = new(PropertyId.Account, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ACKNOWLEDGEMENT_MODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ACKNOWLEDGEMENT_MODE.
    /// </remarks>
    public static readonly PropertyTag AcknowledgementMode = new(PropertyId.AcknowledgementMode, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag AddrtypeA = new(PropertyId.Addrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag AddrtypeW = new(PropertyId.Addrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ALTERNATE_RECIPIENT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ALTERNATE_RECIPIENT.
    /// </remarks>
    public static readonly PropertyTag AlternateRecipient = new(PropertyId.AlternateRecipient, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ALTERNATE_RECIPIENT_ALLOWED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ALTERNATE_RECIPIENT_ALLOWED.
    /// </remarks>
    public static readonly PropertyTag AlternateRecipientAllowed =
        new(PropertyId.AlternateRecipientAllowed, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_ANR.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ANR.
    /// </remarks>
    public static readonly PropertyTag AnrA = new(PropertyId.Anr, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ANR.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ANR.
    /// </remarks>
    public static readonly PropertyTag AnrW = new(PropertyId.Anr, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ASSISTANT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ASSISTANT.
    /// </remarks>
    public static readonly PropertyTag AssistantA = new(PropertyId.Assistant, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ASSISTANT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ASSISTANT.
    /// </remarks>
    public static readonly PropertyTag AssistantW = new(PropertyId.Assistant, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag AssistantTelephoneNumberA =
        new(PropertyId.AssistantTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ASSISTANT_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag AssistantTelephoneNumberW =
        new(PropertyId.AssistantTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ASSOC_CONTENT_COUNT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ASSOC_CONTENT_COUNT.
    /// </remarks>
    public static readonly PropertyTag AssocContentCount = new(PropertyId.AssocContentCount, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ATTACH_ADDITIONAL_INFO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_ADDITIONAL_INFO.
    /// </remarks>
    public static readonly PropertyTag AttachAdditionalInfo = new(PropertyId.AttachAdditionalInfo, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ATTACH_CONTENT_BASE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_CONTENT_BASE.
    /// </remarks>
    public static readonly PropertyTag AttachContentBaseA = new(PropertyId.AttachContentBase, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_CONTENT_BASE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_CONTENT_BASE.
    /// </remarks>
    public static readonly PropertyTag AttachContentBaseW = new(PropertyId.AttachContentBase, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_CONTENT_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_CONTENT_ID.
    /// </remarks>
    public static readonly PropertyTag AttachContentIdA = new(PropertyId.AttachContentId, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_CONTENT_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_CONTENT_ID.
    /// </remarks>
    public static readonly PropertyTag AttachContentIdW = new(PropertyId.AttachContentId, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_CONTENT_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_CONTENT_LOCATION.
    /// </remarks>
    public static readonly PropertyTag AttachContentLocationA =
        new(PropertyId.AttachContentLocation, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_CONTENT_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_CONTENT_LOCATION.
    /// </remarks>
    public static readonly PropertyTag AttachContentLocationW =
        new(PropertyId.AttachContentLocation, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_DATA.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_DATA.
    /// </remarks>
    public static readonly PropertyTag AttachDataBin = new(PropertyId.AttachData, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ATTACH_DATA.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_DATA.
    /// </remarks>
    public static readonly PropertyTag AttachDataObj = new(PropertyId.AttachData, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_ATTACH_DISPOSITION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_DISPOSITION.
    /// </remarks>
    public static readonly PropertyTag AttachDispositionA = new(PropertyId.AttachDisposition, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_DISPOSITION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_DISPOSITION.
    /// </remarks>
    public static readonly PropertyTag AttachDispositionW = new(PropertyId.AttachDisposition, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_ENCODING.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_ENCODING.
    /// </remarks>
    public static readonly PropertyTag AttachEncoding = new(PropertyId.AttachEncoding, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ATTACH_EXTENSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_EXTENSION.
    /// </remarks>
    public static readonly PropertyTag AttachExtensionA = new(PropertyId.AttachExtension, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_EXTENSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_EXTENSION.
    /// </remarks>
    public static readonly PropertyTag AttachExtensionW = new(PropertyId.AttachExtension, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_FILENAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_FILENAME.
    /// </remarks>
    public static readonly PropertyTag AttachFilenameA = new(PropertyId.AttachFilename, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_FILENAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_FILENAME.
    /// </remarks>
    public static readonly PropertyTag AttachFilenameW = new(PropertyId.AttachFilename, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_FLAGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_FLAGS.
    /// </remarks>
    public static readonly PropertyTag AttachFlags = new(PropertyId.AttachFlags, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_ATTACH_LONG_FILENAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_LONG_FILENAME.
    /// </remarks>
    public static readonly PropertyTag AttachLongFilenameA = new(PropertyId.AttachLongFilename, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_LONG_FILENAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_LONG_FILENAME.
    /// </remarks>
    public static readonly PropertyTag AttachLongFilenameW = new(PropertyId.AttachLongFilename, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_LONG_PATHNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_LONG_PATHNAME.
    /// </remarks>
    public static readonly PropertyTag AttachLongPathnameA = new(PropertyId.AttachLongPathname, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_LONG_PATHNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_LONG_PATHNAME.
    /// </remarks>
    public static readonly PropertyTag AttachLongPathnameW = new(PropertyId.AttachLongPathname, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACHMENT_CONTACTPHOTO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACHMENT_CONTACTPHOTO.
    /// </remarks>
    public static readonly PropertyTag AttachmentContactPhoto =
        new(PropertyId.AttachmentContactPhoto, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_ATTACHMENT_FLAGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACHMENT_FLAGS.
    /// </remarks>
    public static readonly PropertyTag AttachmentFlags = new(PropertyId.AttachmentFlags, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ATTACHMENT_HIDDEN.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACHMENT_HIDDEN.
    /// </remarks>
    public static readonly PropertyTag AttachmentHidden = new(PropertyId.AttachmentHidden, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_ATTACHMENT_LINKID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACHMENT_LINKID.
    /// </remarks>
    public static readonly PropertyTag AttachmentLinkId = new(PropertyId.AttachmentLinkId, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ATTACHMENT_X400_PARAMETERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACHMENT_X400_PARAMETERS.
    /// </remarks>
    public static readonly PropertyTag AttachmentX400Parameters =
        new(PropertyId.AttachmentX400Parameters, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ATTACH_METHOD.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_METHOD.
    /// </remarks>
    public static readonly PropertyTag AttachMethod = new(PropertyId.AttachMethod, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ATTACH_MIME_SEQUENCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_MIME_SEQUENCE.
    /// </remarks>
    public static readonly PropertyTag
        AttachMimeSequence = new(PropertyId.AttachMimeSequence, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_ATTACH_MIME_TAG.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_MIME_TAG.
    /// </remarks>
    public static readonly PropertyTag AttachMimeTagA = new(PropertyId.AttachMimeTag, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_MIME_TAG.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_MIME_TAG.
    /// </remarks>
    public static readonly PropertyTag AttachMimeTagW = new(PropertyId.AttachMimeTag, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_NETSCAPE_MAC_INFO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_NETSCAPE_MAC_INFO.
    /// </remarks>
    public static readonly PropertyTag AttachNetscapeMacInfo =
        new(PropertyId.AttachNetscapeMacInfo, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_ATTACH_NUM.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_NUM.
    /// </remarks>
    public static readonly PropertyTag AttachNum = new(PropertyId.AttachNum, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ATTACH_PATHNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_PATHNAME.
    /// </remarks>
    public static readonly PropertyTag AttachPathnameA = new(PropertyId.AttachPathname, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_PATHNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_PATHNAME.
    /// </remarks>
    public static readonly PropertyTag AttachPathnameW = new(PropertyId.AttachPathname, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ATTACH_RENDERING.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_RENDERING.
    /// </remarks>
    public static readonly PropertyTag AttachRendering = new(PropertyId.AttachRendering, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ATTACH_SIZE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_SIZE.
    /// </remarks>
    public static readonly PropertyTag AttachSize = new(PropertyId.AttachSize, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ATTACH_TAG.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_TAG.
    /// </remarks>
    public static readonly PropertyTag AttachTag = new(PropertyId.AttachTag, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ATTACH_TRANSPORT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_TRANSPORT_NAME.
    /// </remarks>
    public static readonly PropertyTag AttachTransportNameA = new(PropertyId.AttachTransportName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ATTACH_TRANSPORT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ATTACH_TRANSPORT_NAME.
    /// </remarks>
    public static readonly PropertyTag AttachTransportNameW = new(PropertyId.AttachTransportName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_AUTHORIZING_USERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AUTHORIZING_USERS.
    /// </remarks>
    public static readonly PropertyTag AuthorizingUsers = new(PropertyId.AuthorizingUsers, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_AUTOFORWARDED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AUTOFORWARDED.
    /// </remarks>
    public static readonly PropertyTag AutoForwarded = new(PropertyId.AutoForwarded, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_AUTOFORWARDING_COMMENT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AUTOFORWARDING_COMMENT.
    /// </remarks>
    public static readonly PropertyTag AutoForwardingCommentA =
        new(PropertyId.AutoForwardingComment, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_AUTOFORWARDING_COMMENT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AUTOFORWARDING_COMMENT.
    /// </remarks>
    public static readonly PropertyTag AutoForwardingCommentW =
        new(PropertyId.AutoForwardingComment, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_AUTORESPONSE_SUPPRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_AUTORESPONSE_SUPPRESS.
    /// </remarks>
    public static readonly PropertyTag AutoResponseSuppress =
        new(PropertyId.AutoResponseSuppress, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag BeeperTelephoneNumberA =
        new(PropertyId.BeeperTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BEEPER_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag BeeperTelephoneNumberW =
        new(PropertyId.BeeperTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BIRTHDAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BIRTHDAY.
    /// </remarks>
    public static readonly PropertyTag Birthday = new(PropertyId.Birthday, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_BODY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY.
    /// </remarks>
    public static readonly PropertyTag BodyA = new(PropertyId.Body, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BODY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY.
    /// </remarks>
    public static readonly PropertyTag BodyW = new(PropertyId.Body, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BODY_CONTENT_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_CONTENT_ID.
    /// </remarks>
    public static readonly PropertyTag BodyContentIdA = new(PropertyId.BodyContentId, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BODY_CONTENT_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_CONTENT_ID.
    /// </remarks>
    public static readonly PropertyTag BodyContentIdW = new(PropertyId.BodyContentId, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BODY_CONTENT_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_CONTENT_LOCATION.
    /// </remarks>
    public static readonly PropertyTag BodyContentLocationA = new(PropertyId.BodyContentLocation, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BODY_CONTENT_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_CONTENT_LOCATION.
    /// </remarks>
    public static readonly PropertyTag BodyContentLocationW = new(PropertyId.BodyContentLocation, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BODY_CRC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_CRC.
    /// </remarks>
    public static readonly PropertyTag BodyCrc = new(PropertyId.BodyCrc, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_BODY_HTML.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_HTML.
    /// </remarks>
    public static readonly PropertyTag BodyHtmlA = new(PropertyId.BodyHtml, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BODY_HTML.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_HTML.
    /// </remarks>
    public static readonly PropertyTag BodyHtmlB = new(PropertyId.BodyHtml, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_BODY_HTML.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BODY_HTML.
    /// </remarks>
    public static readonly PropertyTag BodyHtmlW = new(PropertyId.BodyHtml, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Business2TelephoneNumberA =
        new(PropertyId.Business2TelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Business2TelephoneNumberAMv =
        new(PropertyId.Business2TelephoneNumber, PropertyType.String8, true);

    /// <summary>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Business2TelephoneNumberW =
        new(PropertyId.Business2TelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Business2TelephoneNumberWMv =
        new(PropertyId.Business2TelephoneNumber, PropertyType.Unicode, true);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_CITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_CITY.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressCityA = new(PropertyId.BusinessAddressCity, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_CITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_CITY.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressCityW = new(PropertyId.BusinessAddressCity, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressCountryA =
        new(PropertyId.BusinessAddressCountry, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressCountryW =
        new(PropertyId.BusinessAddressCountry, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressPostalCodeA =
        new(PropertyId.BusinessAddressPostalCode, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressPostalCodeW =
        new(PropertyId.BusinessAddressPostalCode, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressStreetA =
        new(PropertyId.BusinessAddressStreet, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_ADDRESS_POSTAL_STREET.
    /// </remarks>
    public static readonly PropertyTag BusinessAddressStreetW =
        new(PropertyId.BusinessAddressStreet, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_FAX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_FAX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag BusinessFaxNumberA = new(PropertyId.BusinessFaxNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_FAX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_FAX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag BusinessFaxNumberW = new(PropertyId.BusinessFaxNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_HOME_PAGE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_HOME_PAGE.
    /// </remarks>
    public static readonly PropertyTag BusinessHomePageA = new(PropertyId.BusinessHomePage, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_BUSINESS_HOME_PAGE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_BUSINESS_HOME_PAGE.
    /// </remarks>
    public static readonly PropertyTag BusinessHomePageW = new(PropertyId.BusinessHomePage, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag CallbackTelephoneNumberA =
        new(PropertyId.CallbackTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CALLBACK_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag CallbackTelephoneNumberW =
        new(PropertyId.CallbackTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CAR_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CAR_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag CarTelephoneNumberA = new(PropertyId.CarTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_CAR_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CAR_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag CarTelephoneNumberW = new(PropertyId.CarTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CHILDRENS_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CHILDRENS_NAMES.
    /// </remarks>
    public static readonly PropertyTag ChildrensNamesA = new(PropertyId.ChildrensNames, PropertyType.String8, true);

    /// <summary>
    ///     The MAPI property PR_CHILDRENS_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CHILDRENS_NAMES.
    /// </remarks>
    public static readonly PropertyTag ChildrensNamesW = new(PropertyId.ChildrensNames, PropertyType.Unicode, true);

    /// <summary>
    ///     The MAPI property PR_CLIENT_SUBMIT_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CLIENT_SUBMIT_TIME.
    /// </remarks>
    public static readonly PropertyTag ClientSubmitTime = new(PropertyId.ClientSubmitTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_COMMENT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMMENT.
    /// </remarks>
    public static readonly PropertyTag CommentA = new(PropertyId.Comment, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_COMMENT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMMENT.
    /// </remarks>
    public static readonly PropertyTag CommentW = new(PropertyId.Comment, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_COMMON_VIEWS_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMMON_VIEWS_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag CommonViewsEntryId = new(PropertyId.CommonViewsEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag CompanyMainPhoneNumberA =
        new(PropertyId.CompanyMainPhoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMPANY_MAIN_PHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag CompanyMainPhoneNumberW =
        new(PropertyId.CompanyMainPhoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_COMPANY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMPANY_NAME.
    /// </remarks>
    public static readonly PropertyTag CompanyNameA = new(PropertyId.CompanyName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_COMPANY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMPANY_NAME.
    /// </remarks>
    public static readonly PropertyTag CompanyNameW = new(PropertyId.CompanyName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_COMPUTER_NETWORK_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMPUTER_NETWORK_NAME.
    /// </remarks>
    public static readonly PropertyTag ComputerNetworkNameA = new(PropertyId.ComputerNetworkName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_COMPUTER_NETWORK_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COMPUTER_NETWORK_NAME.
    /// </remarks>
    public static readonly PropertyTag ComputerNetworkNameW = new(PropertyId.ComputerNetworkName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CONTACT_ADDRTYPES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTACT_ADDRTYPES.
    /// </remarks>
    public static readonly PropertyTag ContactAddrtypesA = new(PropertyId.ContactAddrtypes, PropertyType.String8, true);

    /// <summary>
    ///     The MAPI property PR_CONTACT_ADDRTYPES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTACT_ADDRTYPES.
    /// </remarks>
    public static readonly PropertyTag ContactAddrtypesW = new(PropertyId.ContactAddrtypes, PropertyType.Unicode, true);

    /// <summary>
    ///     The MAPI property PR_CONTACT_DEFAULT_ADDRESS_INDEX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTACT_DEFAULT_ADDRESS_INDEX.
    /// </remarks>
    public static readonly PropertyTag ContactDefaultAddressIndex =
        new(PropertyId.ContactDefaultAddressIndex, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
    /// </remarks>
    public static readonly PropertyTag ContactEmailAddressesA =
        new(PropertyId.ContactEmailAddresses, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTACT_EMAIL_ADDRESSES.
    /// </remarks>
    public static readonly PropertyTag ContactEmailAddressesW =
        new(PropertyId.ContactEmailAddresses, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CONTACT_ENTRYIDS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTACT_ENTRYIDS.
    /// </remarks>
    public static readonly PropertyTag ContactEntryIds = new(PropertyId.ContactEntryIds, PropertyType.Binary, true);

    /// <summary>
    ///     The MAPI property PR_CONTACT_VERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTACT_VERSION.
    /// </remarks>
    public static readonly PropertyTag ContactVersion = new(PropertyId.ContactVersion, PropertyType.ClassId);

    /// <summary>
    ///     The MAPI property PR_CONTAINER_CLASS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTAINER_CLASS.
    /// </remarks>
    public static readonly PropertyTag ContainerClassA = new(PropertyId.ContainerClass, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_CONTAINER_CLASS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTAINER_CLASS.
    /// </remarks>
    public static readonly PropertyTag ContainerClassW = new(PropertyId.ContainerClass, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CONTAINER_CONTENTS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTAINER_CONTENTS.
    /// </remarks>
    public static readonly PropertyTag ContainerContents = new(PropertyId.ContainerContents, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_CONTAINER_FLAGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTAINER_FLAGS.
    /// </remarks>
    public static readonly PropertyTag ContainerFlags = new(PropertyId.ContainerFlags, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_CONTAINER_HIERARCHY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTAINER_HIERARCHY.
    /// </remarks>
    public static readonly PropertyTag ContainerHierarchy = new(PropertyId.ContainerHierarchy, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_CONTAINER_MODIFY_VERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTAINER_MODIFY_VERSION.
    /// </remarks>
    public static readonly PropertyTag ContainerModifyVersion = new(PropertyId.ContainerModifyVersion, PropertyType.I8);

    /// <summary>
    ///     The MAPI property PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID.
    /// </remarks>
    public static readonly PropertyTag ContentConfidentialityAlgorithmId =
        new(PropertyId.ContentConfidentialityAlgorithmId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONTENT_CORRELATOR.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_CORRELATOR.
    /// </remarks>
    public static readonly PropertyTag ContentCorrelator = new(PropertyId.ContentCorrelator, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONTENT_COUNT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_COUNT.
    /// </remarks>
    public static readonly PropertyTag ContentCount = new(PropertyId.ContentCount, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_CONTENT_IDENTIFIER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_IDENTIFIER.
    /// </remarks>
    public static readonly PropertyTag ContentIdentifierA = new(PropertyId.ContentIdentifier, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_CONTENT_IDENTIFIER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_IDENTIFIER.
    /// </remarks>
    public static readonly PropertyTag ContentIdentifierW = new(PropertyId.ContentIdentifier, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CONTENT_INTEGRITY_CHECK.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_INTEGRITY_CHECK.
    /// </remarks>
    public static readonly PropertyTag ContentIntegrityCheck =
        new(PropertyId.ContentIntegrityCheck, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONTENT_LENGTH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_LENGTH.
    /// </remarks>
    public static readonly PropertyTag ContentLength = new(PropertyId.ContentLength, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_CONTENT_RETURN_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_RETURN_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag ContentReturnRequested =
        new(PropertyId.ContentReturnRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_CONTENTS_SORT_ORDER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENTS_SORT_ORDER.
    /// </remarks>
    public static readonly PropertyTag ContentsSortOrder = new(PropertyId.ContentsSortOrder, PropertyType.Long, true);

    /// <summary>
    ///     The MAPI property PR_CONTENT_UNREAD.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTENT_UNREAD.
    /// </remarks>
    public static readonly PropertyTag ContentUnread = new(PropertyId.ContentUnread, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_CONTROL_FLAGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTROL_FLAGS.
    /// </remarks>
    public static readonly PropertyTag ControlFlags = new(PropertyId.ControlFlags, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_CONTROL_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTROL_ID.
    /// </remarks>
    public static readonly PropertyTag ControlId = new(PropertyId.ControlId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONTROL_STRUCTURE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTROL_STRUCTURE.
    /// </remarks>
    public static readonly PropertyTag ControlStructure = new(PropertyId.ControlStructure, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONTROL_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONTROL_TYPE.
    /// </remarks>
    public static readonly PropertyTag ControlType = new(PropertyId.ControlType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_CONVERSATION_INDEX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERSATION_INDEX.
    /// </remarks>
    public static readonly PropertyTag ConversationIndex = new(PropertyId.ConversationIndex, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONVERSATION_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERSATION_KEY.
    /// </remarks>
    public static readonly PropertyTag ConversationKey = new(PropertyId.ConversationKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONVERSATION_TOPIC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERSATION_TOPIC.
    /// </remarks>
    public static readonly PropertyTag ConversationTopicA = new(PropertyId.ConversationTopic, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_CONVERSATION_TOPIC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERSATION_TOPIC.
    /// </remarks>
    public static readonly PropertyTag ConversationTopicW = new(PropertyId.ConversationTopic, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CONVERSION_EITS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERSION_EITS.
    /// </remarks>
    public static readonly PropertyTag ConversionEits = new(PropertyId.ConversionEits, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CONVERSION_PROHIBITED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERSION_PROHIBITED.
    /// </remarks>
    public static readonly PropertyTag
        ConversionProhibited = new(PropertyId.ConversionProhibited, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_CONVERSION_WITH_LOSS_PROHIBITED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERSION_WITH_LOSS_PROHIBITED.
    /// </remarks>
    public static readonly PropertyTag ConversionWithLossProhibited =
        new(PropertyId.ConversionWithLossProhibited, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_CONVERTED_EITS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CONVERTED_EITS.
    /// </remarks>
    public static readonly PropertyTag ConvertedEits = new(PropertyId.ConvertedEits, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_CORRELATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CORRELATE.
    /// </remarks>
    public static readonly PropertyTag Correlate = new(PropertyId.Correlate, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_CORRELATE_MTSID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CORRELATE_MTSID.
    /// </remarks>
    public static readonly PropertyTag CorrelateMtsid = new(PropertyId.CorrelateMtsid, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag CountryA = new(PropertyId.Country, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag CountryW = new(PropertyId.Country, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_CREATE_TEMPLATES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CREATE_TEMPLATES.
    /// </remarks>
    public static readonly PropertyTag CreateTemplates = new(PropertyId.CreateTemplates, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_CREATION_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CREATION_TIME.
    /// </remarks>
    public static readonly PropertyTag CreationTime = new(PropertyId.CreationTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_CREATION_VERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CREATION_VERSION.
    /// </remarks>
    public static readonly PropertyTag CreationVersion = new(PropertyId.CreationVersion, PropertyType.I8);

    /// <summary>
    ///     The MAPI property PR_CURRENT_VERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CURRENT_VERSION.
    /// </remarks>
    public static readonly PropertyTag CurrentVersion = new(PropertyId.CurrentVersion, PropertyType.I8);

    /// <summary>
    ///     The MAPI property PR_CUSTOMER_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CUSTOMER_ID.
    /// </remarks>
    public static readonly PropertyTag CustomerIdA = new(PropertyId.CustomerId, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_CUSTOMER_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_CUSTOMER_ID.
    /// </remarks>
    public static readonly PropertyTag CustomerIdW = new(PropertyId.CustomerId, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_DEFAULT_PROFILE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEFAULT_PROFILE.
    /// </remarks>
    public static readonly PropertyTag DefaultProfile = new(PropertyId.DefaultProfile, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_DEFAULT_STORE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEFAULT_STORE.
    /// </remarks>
    public static readonly PropertyTag DefaultStore = new(PropertyId.DefaultStore, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_DEFAULT_VIEW_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEFAULT_VIEW_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag DefaultViewEntryId = new(PropertyId.DefaultViewEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_DEF_CREATE_DL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEF_CREATE_DL.
    /// </remarks>
    public static readonly PropertyTag DefCreateDl = new(PropertyId.DefCreateDl, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_DEF_CREATE_MAILUSER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEF_CREATE_MAILUSER.
    /// </remarks>
    public static readonly PropertyTag DefCreateMailuser = new(PropertyId.DefCreateMailuser, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_DEFERRED_DELIVERY_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEFERRED_DELIVERY_TIME.
    /// </remarks>
    public static readonly PropertyTag
        DeferredDeliveryTime = new(PropertyId.DeferredDeliveryTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_DELEGATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DELEGATION.
    /// </remarks>
    public static readonly PropertyTag Delegation = new(PropertyId.Delegation, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_DELETE_AFTER_SUBMIT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DELETE_AFTER_SUBMIT.
    /// </remarks>
    public static readonly PropertyTag DeleteAfterSubmit = new(PropertyId.DeleteAfterSubmit, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_DELIVER_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DELIVER_TIME.
    /// </remarks>
    public static readonly PropertyTag DeliverTime = new(PropertyId.DeliverTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_DELIVERY_POINT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DELIVERY_POINT.
    /// </remarks>
    public static readonly PropertyTag DeliveryPoint = new(PropertyId.DeliveryPoint, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_DELTAX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DELTAX.
    /// </remarks>
    public static readonly PropertyTag Deltax = new(PropertyId.Deltax, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_DELTAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DELTAY.
    /// </remarks>
    public static readonly PropertyTag Deltay = new(PropertyId.Deltay, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_DEPARTMENT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEPARTMENT_NAME.
    /// </remarks>
    public static readonly PropertyTag DepartmentNameA = new(PropertyId.DepartmentName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_DEPARTMENT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEPARTMENT_NAME.
    /// </remarks>
    public static readonly PropertyTag DepartmentNameW = new(PropertyId.DepartmentName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_DEPTH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DEPTH.
    /// </remarks>
    public static readonly PropertyTag Depth = new(PropertyId.Depth, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_DETAILS_TABLE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DETAILS_TABLE.
    /// </remarks>
    public static readonly PropertyTag DetailsTable = new(PropertyId.DetailsTable, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_DISCARD_REASON.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISCARD_REASON.
    /// </remarks>
    public static readonly PropertyTag DiscardReason = new(PropertyId.DiscardReason, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_DISCLOSE_RECIPIENTS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISCLOSE_RECIPIENTS.
    /// </remarks>
    public static readonly PropertyTag DiscloseRecipients = new(PropertyId.DiscloseRecipients, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_DISCLOSURE_OF_RECIPIENTS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISCLOSURE_OF_RECIPIENTS.
    /// </remarks>
    public static readonly PropertyTag DisclosureOfRecipients =
        new(PropertyId.DisclosureOfRecipients, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_DISCRETE_VALUES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISCRETE_VALUES.
    /// </remarks>
    public static readonly PropertyTag DiscreteValues = new(PropertyId.DiscreteValues, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_DISC_VAL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISC_VAL.
    /// </remarks>
    public static readonly PropertyTag DiscVal = new(PropertyId.DiscVal, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_BCC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_BCC.
    /// </remarks>
    public static readonly PropertyTag DisplayBccA = new(PropertyId.DisplayBcc, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_BCC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_BCC.
    /// </remarks>
    public static readonly PropertyTag DisplayBccW = new(PropertyId.DisplayBcc, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_CC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_CC.
    /// </remarks>
    public static readonly PropertyTag DisplayCcA = new(PropertyId.DisplayCc, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_CC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_CC.
    /// </remarks>
    public static readonly PropertyTag DisplayCcW = new(PropertyId.DisplayCc, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag DisplayNameA = new(PropertyId.DisplayName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag DisplayNameW = new(PropertyId.DisplayName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_NAME_PREFIX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_NAME_PREFIX.
    /// </remarks>
    public static readonly PropertyTag DisplayNamePrefixA = new(PropertyId.DisplayNamePrefix, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_NAME_PREFIX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_NAME_PREFIX.
    /// </remarks>
    public static readonly PropertyTag DisplayNamePrefixW = new(PropertyId.DisplayNamePrefix, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_TO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_TO.
    /// </remarks>
    public static readonly PropertyTag DisplayToA = new(PropertyId.DisplayTo, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_TO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_TO.
    /// </remarks>
    public static readonly PropertyTag DisplayToW = new(PropertyId.DisplayTo, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_DISPLAY_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DISPLAY_TYPE.
    /// </remarks>
    public static readonly PropertyTag DisplayType = new(PropertyId.DisplayType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_DL_EXPANSION_HISTORY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DL_EXPANSION_HISTORY.
    /// </remarks>
    public static readonly PropertyTag DlExpansionHistory = new(PropertyId.DlExpansionHistory, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_DL_EXPANSION_PROHIBITED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_DL_EXPANSION_PROHIBITED.
    /// </remarks>
    public static readonly PropertyTag DlExpansionProhibited =
        new(PropertyId.DlExpansionProhibited, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag EmailAddressA = new(PropertyId.EmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag EmailAddressW = new(PropertyId.EmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_END_DATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_END_DATE.
    /// </remarks>
    public static readonly PropertyTag EndDate = new(PropertyId.EndDate, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag EntryId = new(PropertyId.EntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_EXPAND_BEGIN_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EXPAND_BEGIN_TIME.
    /// </remarks>
    public static readonly PropertyTag ExpandBeginTime = new(PropertyId.ExpandBeginTime, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_EXPANDED_BEGIN_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EXPANDED_BEGIN_TIME.
    /// </remarks>
    public static readonly PropertyTag ExpandedBeginTime = new(PropertyId.ExpandedBeginTime, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_EXPANDED_END_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EXPANDED_END_TIME.
    /// </remarks>
    public static readonly PropertyTag ExpandedEndTime = new(PropertyId.ExpandedEndTime, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_EXPAND_END_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EXPAND_END_TIME.
    /// </remarks>
    public static readonly PropertyTag ExpandEndTime = new(PropertyId.ExpandEndTime, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_EXPIRY_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EXPIRY_TIME.
    /// </remarks>
    public static readonly PropertyTag ExpiryTime = new(PropertyId.ExpiryTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_EXPLICIT_CONVERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_EXPLICIT_CONVERSION.
    /// </remarks>
    public static readonly PropertyTag ExplicitConversion = new(PropertyId.ExplicitConversion, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_FILTERING_HOOKS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FILTERING_HOOKS.
    /// </remarks>
    public static readonly PropertyTag FilteringHooks = new(PropertyId.FilteringHooks, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_FINDER_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FINDER_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag FinderEntryId = new(PropertyId.FinderEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_FOLDER_ASSOCIATED_CONTENTS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FOLDER_ASSOCIATED_CONTENTS.
    /// </remarks>
    public static readonly PropertyTag FolderAssociatedContents =
        new(PropertyId.FolderAssociatedContents, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_FOLDER_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FOLDER_TYPE.
    /// </remarks>
    public static readonly PropertyTag FolderType = new(PropertyId.FolderType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_FORM_CATEGORY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_CATEGORY.
    /// </remarks>
    public static readonly PropertyTag FormCategoryA = new(PropertyId.FormCategory, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_FORM_CATEGORY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_CATEGORY.
    /// </remarks>
    public static readonly PropertyTag FormCategoryW = new(PropertyId.FormCategory, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_FORM_CATEGORY_SUB.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_CATEGORY_SUB.
    /// </remarks>
    public static readonly PropertyTag FormCategorySubA = new(PropertyId.FormCategorySub, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_FORM_CATEGORY_SUB.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_CATEGORY_SUB.
    /// </remarks>
    public static readonly PropertyTag FormCategorySubW = new(PropertyId.FormCategorySub, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_FORM_CLSID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_CLSID.
    /// </remarks>
    public static readonly PropertyTag FormClsid = new(PropertyId.FormClsid, PropertyType.ClassId);

    /// <summary>
    ///     The MAPI property PR_FORM_CONTACT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_CONTACT_NAME.
    /// </remarks>
    public static readonly PropertyTag FormContactNameA = new(PropertyId.FormContactName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_FORM_CONTACT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_CONTACT_NAME.
    /// </remarks>
    public static readonly PropertyTag FormContactNameW = new(PropertyId.FormContactName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_FORM_DESIGNER_GUID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_DESIGNER_GUID.
    /// </remarks>
    public static readonly PropertyTag FormDesignerGuid = new(PropertyId.FormDesignerGuid, PropertyType.ClassId);

    /// <summary>
    ///     The MAPI property PR_FORM_DESIGNER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_DESIGNER_NAME.
    /// </remarks>
    public static readonly PropertyTag FormDesignerNameA = new(PropertyId.FormDesignerName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_FORM_DESIGNER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_DESIGNER_NAME.
    /// </remarks>
    public static readonly PropertyTag FormDesignerNameW = new(PropertyId.FormDesignerName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_FORM_HIDDEN.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_HIDDEN.
    /// </remarks>
    public static readonly PropertyTag FormHidden = new(PropertyId.FormHidden, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_FORM_HOST_MAP.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_HOST_MAP.
    /// </remarks>
    public static readonly PropertyTag FormHostMap = new(PropertyId.FormHostMap, PropertyType.Long, true);

    /// <summary>
    ///     The MAPI property PR_FORM_MESSAGE_BEHAVIOR.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_MESSAGE_BEHAVIOR.
    /// </remarks>
    public static readonly PropertyTag FormMessageBehavior = new(PropertyId.FormMessageBehavior, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_FORM_VERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_VERSION.
    /// </remarks>
    public static readonly PropertyTag FormVersionA = new(PropertyId.FormVersion, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_FORM_VERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FORM_VERSION.
    /// </remarks>
    public static readonly PropertyTag FormVersionW = new(PropertyId.FormVersion, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_FTP_SITE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FTP_SITE.
    /// </remarks>
    public static readonly PropertyTag FtpSiteA = new(PropertyId.FtpSite, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_FTP_SITE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_FTP_SITE.
    /// </remarks>
    public static readonly PropertyTag FtpSiteW = new(PropertyId.FtpSite, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_GENDER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_GENDER.
    /// </remarks>
    public static readonly PropertyTag Gender = new(PropertyId.Gender, PropertyType.I2);

    /// <summary>
    ///     The MAPI property PR_GENERATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_GENERATION.
    /// </remarks>
    public static readonly PropertyTag GenerationA = new(PropertyId.Generation, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_GENERATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_GENERATION.
    /// </remarks>
    public static readonly PropertyTag GenerationW = new(PropertyId.Generation, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_GIVEN_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_GIVEN_NAME.
    /// </remarks>
    public static readonly PropertyTag GivenNameA = new(PropertyId.GivenName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_GIVEN_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_GIVEN_NAME.
    /// </remarks>
    public static readonly PropertyTag GivenNameW = new(PropertyId.GivenName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_GOVERNMENT_ID_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_GOVERNMENT_ID_NUMBER.
    /// </remarks>
    public static readonly PropertyTag GovernmentIdNumberA = new(PropertyId.GovernmentIdNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_GOVERNMENT_ID_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_GOVERNMENT_ID_NUMBER.
    /// </remarks>
    public static readonly PropertyTag GovernmentIdNumberW = new(PropertyId.GovernmentIdNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HASATTACH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HASATTACH.
    /// </remarks>
    public static readonly PropertyTag Hasattach = new(PropertyId.Hasattach, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_HEADER_FOLDER_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HEADER_FOLDER_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag HeaderFolderEntryId = new(PropertyId.HeaderFolderEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_HOBBIES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOBBIES.
    /// </remarks>
    public static readonly PropertyTag HobbiesA = new(PropertyId.Hobbies, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOBBIES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOBBIES.
    /// </remarks>
    public static readonly PropertyTag HobbiesW = new(PropertyId.Hobbies, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Home2TelephoneNumberA =
        new(PropertyId.Home2TelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Home2TelephoneNumberAMv =
        new(PropertyId.Home2TelephoneNumber, PropertyType.String8, true);

    /// <summary>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Home2TelephoneNumberW =
        new(PropertyId.Home2TelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Home2TelephoneNumberWMv =
        new(PropertyId.Home2TelephoneNumber, PropertyType.Unicode, true);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_CITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_CITY.
    /// </remarks>
    public static readonly PropertyTag HomeAddressCityA = new(PropertyId.HomeAddressCity, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_CITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_CITY.
    /// </remarks>
    public static readonly PropertyTag HomeAddressCityW = new(PropertyId.HomeAddressCity, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag HomeAddressCountryA = new(PropertyId.HomeAddressCountry, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag HomeAddressCountryW = new(PropertyId.HomeAddressCountry, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag HomeAddressPostalCodeA =
        new(PropertyId.HomeAddressPostalCode, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag HomeAddressPostalCodeW =
        new(PropertyId.HomeAddressPostalCode, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
    /// </remarks>
    public static readonly PropertyTag HomeAddressPostOfficeBoxA =
        new(PropertyId.HomeAddressPostOfficeBox, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_POST_OFFICE_BOX.
    /// </remarks>
    public static readonly PropertyTag HomeAddressPostOfficeBoxW =
        new(PropertyId.HomeAddressPostOfficeBox, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
    /// </remarks>
    public static readonly PropertyTag HomeAddressStateOrProvinceA =
        new(PropertyId.HomeAddressStateOrProvince, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_STATE_OR_PROVINCE.
    /// </remarks>
    public static readonly PropertyTag HomeAddressStateOrProvinceW =
        new(PropertyId.HomeAddressStateOrProvince, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_STREET.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_STREET.
    /// </remarks>
    public static readonly PropertyTag HomeAddressStreetA = new(PropertyId.HomeAddressStreet, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_ADDRESS_STREET.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_ADDRESS_STREET.
    /// </remarks>
    public static readonly PropertyTag HomeAddressStreetW = new(PropertyId.HomeAddressStreet, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME_FAX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_FAX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag HomeFaxNumberA = new(PropertyId.HomeFaxNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_FAX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_FAX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag HomeFaxNumberW = new(PropertyId.HomeFaxNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_HOME_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag HomeTelephoneNumberA = new(PropertyId.HomeTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_HOME_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_HOME_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag HomeTelephoneNumberW = new(PropertyId.HomeTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ICON.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ICON.
    /// </remarks>
    public static readonly PropertyTag Icon = new(PropertyId.Icon, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IDENTITY_DISPLAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IDENTITY_DISPLAY.
    /// </remarks>
    public static readonly PropertyTag IdentityDisplayA = new(PropertyId.IdentityDisplay, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_IDENTITY_DISPLAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IDENTITY_DISPLAY.
    /// </remarks>
    public static readonly PropertyTag IdentityDisplayW = new(PropertyId.IdentityDisplay, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_IDENTITY_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IDENTITY_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag IdentityEntryId = new(PropertyId.IdentityEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IDENTITY_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IDENTITY_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag IdentitySearchKey = new(PropertyId.IdentitySearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IMPLICIT_CONVERSION_PROHIBITED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IMPLICIT_CONVERSION_PROHIBITED.
    /// </remarks>
    public static readonly PropertyTag ImplicitConversionProhibited =
        new(PropertyId.ImplicitConversionProhibited, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_IMPORTANCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IMPORTANCE.
    /// </remarks>
    public static readonly PropertyTag Importance = new(PropertyId.Importance, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_INCOMPLETE_COPY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INCOMPLETE_COPY.
    /// </remarks>
    public static readonly PropertyTag IncompleteCopy = new(PropertyId.IncompleteCopy, PropertyType.Boolean);

    /// <summary>
    ///     The Internet mail override charset.
    /// </summary>
    /// <remarks>
    ///     The Internet mail override charset.
    /// </remarks>
    public static readonly PropertyTag INetMailOverrideCharset =
        new(PropertyId.INetMailOverrideCharset, PropertyType.Unspecified);

    /// <summary>
    ///     The Internet mail override format.
    /// </summary>
    /// <remarks>
    ///     The Internet mail override format.
    /// </remarks>
    public static readonly PropertyTag INetMailOverrideFormat =
        new(PropertyId.INetMailOverrideFormat, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_INITIAL_DETAILS_PANE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INITIAL_DETAILS_PANE.
    /// </remarks>
    public static readonly PropertyTag InitialDetailsPane = new(PropertyId.InitialDetailsPane, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_INITIALS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INITIALS.
    /// </remarks>
    public static readonly PropertyTag InitialsA = new(PropertyId.Initials, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INITIALS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INITIALS.
    /// </remarks>
    public static readonly PropertyTag InitialsW = new(PropertyId.Initials, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_IN_REPLY_TO_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IN_REPLY_TO_ID.
    /// </remarks>
    public static readonly PropertyTag InReplyToIdA = new(PropertyId.InReplyToId, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_IN_REPLY_TO_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IN_REPLY_TO_ID.
    /// </remarks>
    public static readonly PropertyTag InReplyToIdW = new(PropertyId.InReplyToId, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INSTANCE_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INSTANCE_KEY.
    /// </remarks>
    public static readonly PropertyTag InstanceKey = new(PropertyId.InstanceKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_INTERNET_APPROVED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_APPROVED.
    /// </remarks>
    public static readonly PropertyTag InternetApprovedA = new(PropertyId.InternetApproved, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_APPROVED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_APPROVED.
    /// </remarks>
    public static readonly PropertyTag InternetApprovedW = new(PropertyId.InternetApproved, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_ARTICLE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_ARTICLE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag InternetArticleNumber = new(PropertyId.InternetArticleNumber, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_INTERNET_CONTROL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_CONTROL.
    /// </remarks>
    public static readonly PropertyTag InternetControlA = new(PropertyId.InternetControl, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_CONTROL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_CONTROL.
    /// </remarks>
    public static readonly PropertyTag InternetControlW = new(PropertyId.InternetControl, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_CPID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_CPID.
    /// </remarks>
    public static readonly PropertyTag InternetCPID = new(PropertyId.InternetCPID, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_INTERNET_DISTRIBUTION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_DISTRIBUTION.
    /// </remarks>
    public static readonly PropertyTag InternetDistributionA =
        new(PropertyId.InternetDistribution, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_DISTRIBUTION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_DISTRIBUTION.
    /// </remarks>
    public static readonly PropertyTag InternetDistributionW =
        new(PropertyId.InternetDistribution, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_FOLLOWUP_TO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_FOLLOWUP_TO.
    /// </remarks>
    public static readonly PropertyTag InternetFollowupToA = new(PropertyId.InternetFollowupTo, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_FOLLOWUP_TO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_FOLLOWUP_TO.
    /// </remarks>
    public static readonly PropertyTag InternetFollowupToW = new(PropertyId.InternetFollowupTo, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_LINES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_LINES.
    /// </remarks>
    public static readonly PropertyTag InternetLines = new(PropertyId.InternetLines, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_INTERNET_MESSAGE_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_MESSAGE_ID.
    /// </remarks>
    public static readonly PropertyTag InternetMessageIdA = new(PropertyId.InternetMessageId, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_MESSAGE_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_MESSAGE_ID.
    /// </remarks>
    public static readonly PropertyTag InternetMessageIdW = new(PropertyId.InternetMessageId, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_NEWSGROUPS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_NEWSGROUPS.
    /// </remarks>
    public static readonly PropertyTag InternetNewsgroupsA = new(PropertyId.InternetNewsgroups, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_NEWSGROUPS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_NEWSGROUPS.
    /// </remarks>
    public static readonly PropertyTag InternetNewsgroupsW = new(PropertyId.InternetNewsgroups, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_NNTP_PATH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_NNTP_PATH.
    /// </remarks>
    public static readonly PropertyTag InternetNntpPathA = new(PropertyId.InternetNntpPath, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_NNTP_PATH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_NNTP_PATH.
    /// </remarks>
    public static readonly PropertyTag InternetNntpPathW = new(PropertyId.InternetNntpPath, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_ORGANIZATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_ORGANIZATION.
    /// </remarks>
    public static readonly PropertyTag InternetOrganizationA =
        new(PropertyId.InternetOrganization, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_ORGANIZATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_ORGANIZATION.
    /// </remarks>
    public static readonly PropertyTag InternetOrganizationW =
        new(PropertyId.InternetOrganization, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_PRECEDENCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_PRECEDENCE.
    /// </remarks>
    public static readonly PropertyTag InternetPrecedenceA = new(PropertyId.InternetPrecedence, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_PRECEDENCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_PRECEDENCE.
    /// </remarks>
    public static readonly PropertyTag InternetPrecedenceW = new(PropertyId.InternetPrecedence, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_INTERNET_REFERENCES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_REFERENCES.
    /// </remarks>
    public static readonly PropertyTag InternetReferencesA = new(PropertyId.InternetReferences, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_INTERNET_REFERENCES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_INTERNET_REFERENCES.
    /// </remarks>
    public static readonly PropertyTag InternetReferencesW = new(PropertyId.InternetReferences, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_IPM_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_ID.
    /// </remarks>
    public static readonly PropertyTag IpmId = new(PropertyId.IpmId, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_IPM_OUTBOX_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_OUTBOX_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag IpmOutboxEntryId = new(PropertyId.IpmOutboxEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IPM_OUTBOX_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_OUTBOX_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag IpmOutboxSearchKey = new(PropertyId.IpmOutboxSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IPM_RETURN_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_RETURN_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag IpmReturnRequested = new(PropertyId.IpmReturnRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_IPM_SENTMAIL_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_SENTMAIL_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag IpmSentmailEntryId = new(PropertyId.IpmSentmailEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IPM_SENTMAIL_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_SENTMAIL_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag IpmSentmailSearchKey = new(PropertyId.IpmSentmailSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IPM_SUBTREE_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_SUBTREE_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag IpmSubtreeEntryId = new(PropertyId.IpmSubtreeEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IPM_SUBTREE_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_SUBTREE_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag IpmSubtreeSearchKey = new(PropertyId.IpmSubtreeSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IPM_WASTEBASKET_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_WASTEBASKET_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag IpmWastebasketEntryId =
        new(PropertyId.IpmWastebasketEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_IPM_WASTEBASKET_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_IPM_WASTEBASKET_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag IpmWastebasketSearchKey =
        new(PropertyId.IpmWastebasketSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ISDN_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ISDN_NUMBER.
    /// </remarks>
    public static readonly PropertyTag IsdnNumberA = new(PropertyId.IsdnNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ISDN_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ISDN_NUMBER.
    /// </remarks>
    public static readonly PropertyTag IsdnNumberW = new(PropertyId.IsdnNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_KEYWORD.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_KEYWORD.
    /// </remarks>
    public static readonly PropertyTag KeywordA = new(PropertyId.Keyword, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_KEYWORD.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_KEYWORD.
    /// </remarks>
    public static readonly PropertyTag KeywordW = new(PropertyId.Keyword, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LANGUAGE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LANGUAGE.
    /// </remarks>
    public static readonly PropertyTag LanguageA = new(PropertyId.Language, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LANGUAGE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LANGUAGE.
    /// </remarks>
    public static readonly PropertyTag LanguageW = new(PropertyId.Language, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LANGUAGES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LANGUAGES.
    /// </remarks>
    public static readonly PropertyTag LanguagesA = new(PropertyId.Languages, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LANGUAGES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LANGUAGES.
    /// </remarks>
    public static readonly PropertyTag LanguagesW = new(PropertyId.Languages, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LAST_MODIFICATION_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LAST_MODIFICATION_TIME.
    /// </remarks>
    public static readonly PropertyTag
        LastModificationTime = new(PropertyId.LastModificationTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_LAST_MODIFIER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LAST_MODIFIER_NAME.
    /// </remarks>
    public static readonly PropertyTag LastModifierNameA = new(PropertyId.LastModifierName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LAST_MODIFIER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LAST_MODIFIER_NAME.
    /// </remarks>
    public static readonly PropertyTag LastModifierNameW = new(PropertyId.LastModifierName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LATEST_DELIVERY_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LATEST_DELIVERY_TIME.
    /// </remarks>
    public static readonly PropertyTag LatestDeliveryTime = new(PropertyId.LatestDeliveryTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_LIST_HELP.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LIST_HELP.
    /// </remarks>
    public static readonly PropertyTag ListHelpA = new(PropertyId.ListHelp, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LIST_HELP.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LIST_HELP.
    /// </remarks>
    public static readonly PropertyTag ListHelpW = new(PropertyId.ListHelp, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LIST_SUBSCRIBE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LIST_SUBSCRIBE.
    /// </remarks>
    public static readonly PropertyTag ListSubscribeA = new(PropertyId.ListSubscribe, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LIST_SUBSCRIBE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LIST_SUBSCRIBE.
    /// </remarks>
    public static readonly PropertyTag ListSubscribeW = new(PropertyId.ListSubscribe, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LIST_UNSUBSCRIBE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LIST_UNSUBSCRIBE.
    /// </remarks>
    public static readonly PropertyTag ListUnsubscribeA = new(PropertyId.ListUnsubscribe, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LIST_UNSUBSCRIBE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LIST_UNSUBSCRIBE.
    /// </remarks>
    public static readonly PropertyTag ListUnsubscribeW = new(PropertyId.ListUnsubscribe, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LOCALITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCALITY.
    /// </remarks>
    public static readonly PropertyTag LocalityA = new(PropertyId.Locality, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LOCALITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCALITY.
    /// </remarks>
    public static readonly PropertyTag LocalityW = new(PropertyId.Locality, PropertyType.Unicode);

    //public static readonly PropertyTag LocallyDelivered = new PropertyTag (TnefPropertyId.LocallyDelivered, TnefPropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCATION.
    /// </remarks>
    public static readonly PropertyTag LocationA = new(PropertyId.Location, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCATION.
    /// </remarks>
    public static readonly PropertyTag LocationW = new(PropertyId.Location, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_LOCK_BRANCH_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_BRANCH_ID.
    /// </remarks>
    public static readonly PropertyTag LockBranchId = new(PropertyId.LockBranchId, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_DEPTH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_DEPTH.
    /// </remarks>
    public static readonly PropertyTag LockDepth = new(PropertyId.LockDepth, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_ENLISTMENT_CONTEXT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_ENLISTMENT_CONTEXT.
    /// </remarks>
    public static readonly PropertyTag LockEnlistmentContext =
        new(PropertyId.LockEnlistmentContext, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_EXPIRY_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_EXPIRY_TIME.
    /// </remarks>
    public static readonly PropertyTag LockExpiryTime = new(PropertyId.LockExpiryTime, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_PERSISTENT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_PERSISTENT.
    /// </remarks>
    public static readonly PropertyTag LockPersistent = new(PropertyId.LockPersistent, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_RESOURCE_DID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_RESOURCE_DID.
    /// </remarks>
    public static readonly PropertyTag LockResourceDid = new(PropertyId.LockResourceDid, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_RESOURCE_FID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_RESOURCE_FID.
    /// </remarks>
    public static readonly PropertyTag LockResourceFid = new(PropertyId.LockResourceFid, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_RESOURCE_MID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_RESOURCE_MID.
    /// </remarks>
    public static readonly PropertyTag LockResourceMid = new(PropertyId.LockResourceMid, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_SCOPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_SCOPE.
    /// </remarks>
    public static readonly PropertyTag LockScope = new(PropertyId.LockScope, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_TIMEOUT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_TIMEOUT.
    /// </remarks>
    public static readonly PropertyTag LockTimeout = new(PropertyId.LockTimeout, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_LOCK_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_LOCK_TYPE.
    /// </remarks>
    public static readonly PropertyTag LockType = new(PropertyId.LockType, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_MAIL_PERMISSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MAIL_PERMISSION.
    /// </remarks>
    public static readonly PropertyTag MailPermission = new(PropertyId.MailPermission, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_MANAGER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MANAGER_NAME.
    /// </remarks>
    public static readonly PropertyTag ManagerNameA = new(PropertyId.ManagerName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_MANAGER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MANAGER_NAME.
    /// </remarks>
    public static readonly PropertyTag ManagerNameW = new(PropertyId.ManagerName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_MAPPING_SIGNATURE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MAPPING_SIGNATURE.
    /// </remarks>
    public static readonly PropertyTag MappingSignature = new(PropertyId.MappingSignature, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_MDB_PROVIDER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MDB_PROVIDER.
    /// </remarks>
    public static readonly PropertyTag MdbProvider = new(PropertyId.MdbProvider, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_ATTACHMENTS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_ATTACHMENTS.
    /// </remarks>
    public static readonly PropertyTag MessageAttachments = new(PropertyId.MessageAttachments, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_CC_ME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_CC_ME.
    /// </remarks>
    public static readonly PropertyTag MessageCcMe = new(PropertyId.MessageCcMe, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_CLASS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_CLASS.
    /// </remarks>
    public static readonly PropertyTag MessageClassA = new(PropertyId.MessageClass, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_CLASS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_CLASS.
    /// </remarks>
    public static readonly PropertyTag MessageClassW = new(PropertyId.MessageClass, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_CODEPAGE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_CODEPAGE.
    /// </remarks>
    public static readonly PropertyTag MessageCodepage = new(PropertyId.MessageCodepage, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_DELIVERY_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_DELIVERY_ID.
    /// </remarks>
    public static readonly PropertyTag MessageDeliveryId = new(PropertyId.MessageDeliveryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_DELIVERY_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_DELIVERY_TIME.
    /// </remarks>
    public static readonly PropertyTag MessageDeliveryTime = new(PropertyId.MessageDeliveryTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_DOWNLOAD_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_DOWNLOAD_TIME.
    /// </remarks>
    public static readonly PropertyTag MessageDownloadTime = new(PropertyId.MessageDownloadTime, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_FLAGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_FLAGS.
    /// </remarks>
    public static readonly PropertyTag MessageFlags = new(PropertyId.MessageFlags, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_RECIPIENTS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_RECIPIENTS.
    /// </remarks>
    public static readonly PropertyTag MessageRecipients = new(PropertyId.MessageRecipients, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_RECIP_ME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_RECIP_ME.
    /// </remarks>
    public static readonly PropertyTag MessageRecipMe = new(PropertyId.MessageRecipMe, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_SECURITY_LABEL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_SECURITY_LABEL.
    /// </remarks>
    public static readonly PropertyTag MessageSecurityLabel = new(PropertyId.MessageSecurityLabel, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_SIZE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_SIZE.
    /// </remarks>
    public static readonly PropertyTag MessageSize = new(PropertyId.MessageSize, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_SUBMISSION_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_SUBMISSION_ID.
    /// </remarks>
    public static readonly PropertyTag MessageSubmissionId = new(PropertyId.MessageSubmissionId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_TOKEN.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_TOKEN.
    /// </remarks>
    public static readonly PropertyTag MessageToken = new(PropertyId.MessageToken, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_MESSAGE_TO_ME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MESSAGE_TO_ME.
    /// </remarks>
    public static readonly PropertyTag MessageToMe = new(PropertyId.MessageToMe, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_MHS_COMMON_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MHS_COMMON_NAME.
    /// </remarks>
    public static readonly PropertyTag MhsCommonNameA = new(PropertyId.MhsCommonName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_MHS_COMMON_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MHS_COMMON_NAME.
    /// </remarks>
    public static readonly PropertyTag MhsCommonNameW = new(PropertyId.MhsCommonName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_MIDDLE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MIDDLE_NAME.
    /// </remarks>
    public static readonly PropertyTag MiddleNameA = new(PropertyId.MiddleName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_MIDDLE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MIDDLE_NAME.
    /// </remarks>
    public static readonly PropertyTag MiddleNameW = new(PropertyId.MiddleName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_MINI_ICON.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MINI_ICON.
    /// </remarks>
    public static readonly PropertyTag MiniIcon = new(PropertyId.MiniIcon, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag MobileTelephoneNumberA =
        new(PropertyId.MobileTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MOBILE_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag MobileTelephoneNumberW =
        new(PropertyId.MobileTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_MODIFY_VERSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MODIFY_VERSION.
    /// </remarks>
    public static readonly PropertyTag ModifyVersion = new(PropertyId.ModifyVersion, PropertyType.I8);

    /// <summary>
    ///     The MAPI property PR_MSG_STATUS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_MSG_STATUS.
    /// </remarks>
    public static readonly PropertyTag MsgStatus = new(PropertyId.MsgStatus, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_NDR_DIAG_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NDR_DIAG_CODE.
    /// </remarks>
    public static readonly PropertyTag NdrDiagCode = new(PropertyId.NdrDiagCode, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_NDR_REASON_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NDR_REASON_CODE.
    /// </remarks>
    public static readonly PropertyTag NdrReasonCode = new(PropertyId.NdrReasonCode, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_NDR_STATUS_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NDR_STATUS_CODE.
    /// </remarks>
    public static readonly PropertyTag NdrStatusCode = new(PropertyId.NdrStatusCode, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_NEWSGROUP_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NEWSGROUP_NAME.
    /// </remarks>
    public static readonly PropertyTag NewsgroupNameA = new(PropertyId.NewsgroupName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_NEWSGROUP_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NEWSGROUP_NAME.
    /// </remarks>
    public static readonly PropertyTag NewsgroupNameW = new(PropertyId.NewsgroupName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_NICKNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NICKNAME.
    /// </remarks>
    public static readonly PropertyTag NicknameA = new(PropertyId.Nickname, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_NICKNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NICKNAME.
    /// </remarks>
    public static readonly PropertyTag NicknameW = new(PropertyId.Nickname, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_NNTP_XREF.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NNTP_XREF.
    /// </remarks>
    public static readonly PropertyTag NntpXrefA = new(PropertyId.NntpXref, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_NNTP_XREF.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NNTP_XREF.
    /// </remarks>
    public static readonly PropertyTag NntpXrefW = new(PropertyId.NntpXref, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_NON_RECEIPT_NOTIFICATION_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NON_RECEIPT_NOTIFICATION_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag NonReceiptNotificationRequested =
        new(PropertyId.NonReceiptNotificationRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_NON_RECEIPT_REASON.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NON_RECEIPT_REASON.
    /// </remarks>
    public static readonly PropertyTag NonReceiptReason = new(PropertyId.NonReceiptReason, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_NORMALIZED_SUBJECT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NORMALIZED_SUBJECT.
    /// </remarks>
    public static readonly PropertyTag NormalizedSubjectA = new(PropertyId.NormalizedSubject, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_NORMALIZED_SUBJECT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NORMALIZED_SUBJECT.
    /// </remarks>
    public static readonly PropertyTag NormalizedSubjectW = new(PropertyId.NormalizedSubject, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_NT_SECURITY_DESCRIPTOR.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NT_SECURITY_DESCRIPTOR.
    /// </remarks>
    public static readonly PropertyTag NtSecurityDescriptor =
        new(PropertyId.NtSecurityDescriptor, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_NULL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_NULL.
    /// </remarks>
    public static readonly PropertyTag Null = new(PropertyId.Null, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_OBJECT_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OBJECT_TYPE.
    /// </remarks>
    public static readonly PropertyTag ObjectType = new(PropertyId.ObjectType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_OBSOLETE_IPMS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OBSOLETE_IPMS.
    /// </remarks>
    public static readonly PropertyTag ObsoletedIpms = new(PropertyId.ObsoletedIpms, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Office2TelephoneNumberA =
        new(PropertyId.Office2TelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OFFICE2_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag Office2TelephoneNumberW =
        new(PropertyId.Office2TelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OFFICE_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OFFICE_LOCATION.
    /// </remarks>
    public static readonly PropertyTag OfficeLocationA = new(PropertyId.OfficeLocation, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OFFICE_LOCATION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OFFICE_LOCATION.
    /// </remarks>
    public static readonly PropertyTag OfficeLocationW = new(PropertyId.OfficeLocation, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag OfficeTelephoneNumberA =
        new(PropertyId.OfficeTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OFFICE_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag OfficeTelephoneNumberW =
        new(PropertyId.OfficeTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OOF_REPLY_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OOF_REPLY_TYPE.
    /// </remarks>
    public static readonly PropertyTag OofReplyType = new(PropertyId.OofReplyType, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
    /// </remarks>
    public static readonly PropertyTag OrganizationalIdNumberA =
        new(PropertyId.OrganizationalIdNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORGANIZATIONAL_ID_NUMBER.
    /// </remarks>
    public static readonly PropertyTag OrganizationalIdNumberW =
        new(PropertyId.OrganizationalIdNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIG_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIG_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag OrigEntryId = new(PropertyId.OrigEntryId, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorAddrtypeA =
        new(PropertyId.OriginalAuthorAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorAddrtypeW =
        new(PropertyId.OriginalAuthorAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorEmailAddressA =
        new(PropertyId.OriginalAuthorEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorEmailAddressW =
        new(PropertyId.OriginalAuthorEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorEntryId =
        new(PropertyId.OriginalAuthorEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorNameA = new(PropertyId.OriginalAuthorName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorNameW = new(PropertyId.OriginalAuthorName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_AUTHOR_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag OriginalAuthorSearchKey =
        new(PropertyId.OriginalAuthorSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DELIVERY_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DELIVERY_TIME.
    /// </remarks>
    public static readonly PropertyTag
        OriginalDeliveryTime = new(PropertyId.OriginalDeliveryTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_BCC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_BCC.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayBccA = new(PropertyId.OriginalDisplayBcc, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_BCC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_BCC.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayBccW = new(PropertyId.OriginalDisplayBcc, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_CC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_CC.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayCcA = new(PropertyId.OriginalDisplayCc, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_CC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_CC.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayCcW = new(PropertyId.OriginalDisplayCc, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayNameA = new(PropertyId.OriginalDisplayName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayNameW = new(PropertyId.OriginalDisplayName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_TO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_TO.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayToA = new(PropertyId.OriginalDisplayTo, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_TO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_DISPLAY_TO.
    /// </remarks>
    public static readonly PropertyTag OriginalDisplayToW = new(PropertyId.OriginalDisplayTo, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_EITS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_EITS.
    /// </remarks>
    public static readonly PropertyTag OriginalEits = new(PropertyId.OriginalEits, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag OriginalEntryId = new(PropertyId.OriginalEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginallyIntendedRecipAddrtypeA =
        new(PropertyId.OriginallyIntendedRecipAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginallyIntendedRecipAddrtypeW =
        new(PropertyId.OriginallyIntendedRecipAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginallyIntendedRecipEmailAddressA =
        new(PropertyId.OriginallyIntendedRecipEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginallyIntendedRecipEmailAddressW =
        new(PropertyId.OriginallyIntendedRecipEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIP_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag OriginallyIntendedRecipEntryId =
        new(PropertyId.OriginallyIntendedRecipEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIPIENT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINALLY_INTENDED_RECIPIENT_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginallyIntendedRecipientName =
        new(PropertyId.OriginallyIntendedRecipientName, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag OriginalSearchKey = new(PropertyId.OriginalSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderAddrtypeA =
        new(PropertyId.OriginalSenderAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderAddrtypeW =
        new(PropertyId.OriginalSenderAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderEmailAddressA =
        new(PropertyId.OriginalSenderEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderEmailAddressW =
        new(PropertyId.OriginalSenderEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderEntryId =
        new(PropertyId.OriginalSenderEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderNameA = new(PropertyId.OriginalSenderName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderNameW = new(PropertyId.OriginalSenderName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENDER_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENDER_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag OriginalSenderSearchKey =
        new(PropertyId.OriginalSenderSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENSITIVITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENSITIVITY.
    /// </remarks>
    public static readonly PropertyTag OriginalSensitivity = new(PropertyId.OriginalSensitivity, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingAddrtypeA =
        new(PropertyId.OriginalSentRepresentingAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingAddrtypeW =
        new(PropertyId.OriginalSentRepresentingAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingEmailAddressA =
        new(PropertyId.OriginalSentRepresentingEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingEmailAddressW =
        new(PropertyId.OriginalSentRepresentingEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingEntryId =
        new(PropertyId.OriginalSentRepresentingEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingNameA =
        new(PropertyId.OriginalSentRepresentingName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_NAME.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingNameW =
        new(PropertyId.OriginalSentRepresentingName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag OriginalSentRepresentingSearchKey =
        new(PropertyId.OriginalSentRepresentingSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SUBJECT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SUBJECT.
    /// </remarks>
    public static readonly PropertyTag OriginalSubjectA = new(PropertyId.OriginalSubject, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SUBJECT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SUBJECT.
    /// </remarks>
    public static readonly PropertyTag OriginalSubjectW = new(PropertyId.OriginalSubject, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_ORIGINAL_SUBMIT_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINAL_SUBMIT_TIME.
    /// </remarks>
    public static readonly PropertyTag OriginalSubmitTime = new(PropertyId.OriginalSubmitTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_ORIGINATING_MTA_CERTIFICATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINATING_MTA_CERTIFICATE.
    /// </remarks>
    public static readonly PropertyTag OriginatingMtaCertificate =
        new(PropertyId.OriginatingMtaCertificate, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY.
    /// </remarks>
    public static readonly PropertyTag OriginatorAndDlExpansionHistory =
        new(PropertyId.OriginatorAndDlExpansionHistory, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINATOR_CERTIFICATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINATOR_CERTIFICATE.
    /// </remarks>
    public static readonly PropertyTag OriginatorCertificate =
        new(PropertyId.OriginatorCertificate, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag OriginatorDeliveryReportRequested =
        new(PropertyId.OriginatorDeliveryReportRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag OriginatorNonDeliveryReportRequested =
        new(PropertyId.OriginatorNonDeliveryReportRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT.
    /// </remarks>
    public static readonly PropertyTag OriginatorRequestedAlternateRecipient =
        new(PropertyId.OriginatorRequestedAlternateRecipient, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGINATOR_RETURN_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGINATOR_RETURN_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag OriginatorReturnAddress =
        new(PropertyId.OriginatorReturnAddress, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIGIN_CHECK.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIGIN_CHECK.
    /// </remarks>
    public static readonly PropertyTag OriginCheck = new(PropertyId.OriginCheck, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_ORIG_MESSAGE_CLASS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIG_MESSAGE_CLASS.
    /// </remarks>
    public static readonly PropertyTag OrigMessageClassA = new(PropertyId.OrigMessageClass, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_ORIG_MESSAGE_CLASS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ORIG_MESSAGE_CLASS.
    /// </remarks>
    public static readonly PropertyTag OrigMessageClassW = new(PropertyId.OrigMessageClass, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_CITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_CITY.
    /// </remarks>
    public static readonly PropertyTag OtherAddressCityA = new(PropertyId.OtherAddressCity, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_CITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_CITY.
    /// </remarks>
    public static readonly PropertyTag OtherAddressCityW = new(PropertyId.OtherAddressCity, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag OtherAddressCountryA = new(PropertyId.OtherAddressCountry, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_COUNTRY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_COUNTRY.
    /// </remarks>
    public static readonly PropertyTag OtherAddressCountryW = new(PropertyId.OtherAddressCountry, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag OtherAddressPostalCodeA =
        new(PropertyId.OtherAddressPostalCode, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag OtherAddressPostalCodeW =
        new(PropertyId.OtherAddressPostalCode, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
    /// </remarks>
    public static readonly PropertyTag OtherAddressPostOfficeBoxA =
        new(PropertyId.OtherAddressPostOfficeBox, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_POST_OFFICE_BOX.
    /// </remarks>
    public static readonly PropertyTag OtherAddressPostOfficeBoxW =
        new(PropertyId.OtherAddressPostOfficeBox, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
    /// </remarks>
    public static readonly PropertyTag OtherAddressStateOrProvinceA =
        new(PropertyId.OtherAddressStateOrProvince, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_STATE_OR_PROVINCE.
    /// </remarks>
    public static readonly PropertyTag OtherAddressStateOrProvinceW =
        new(PropertyId.OtherAddressStateOrProvince, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_STREET.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_STREET.
    /// </remarks>
    public static readonly PropertyTag OtherAddressStreetA = new(PropertyId.OtherAddressStreet, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OTHER_ADDRESS_STREET.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_ADDRESS_STREET.
    /// </remarks>
    public static readonly PropertyTag OtherAddressStreetW = new(PropertyId.OtherAddressStreet, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OTHER_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag OtherTelephoneNumberA =
        new(PropertyId.OtherTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_OTHER_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OTHER_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag OtherTelephoneNumberW =
        new(PropertyId.OtherTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_OWNER_APPT_ID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OWNER_APPT_ID.
    /// </remarks>
    public static readonly PropertyTag OwnerApptId = new(PropertyId.OwnerApptId, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_OWN_STORE_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_OWN_STORE_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag OwnStoreEntryId = new(PropertyId.OwnStoreEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PAGER_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PAGER_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag PagerTelephoneNumberA =
        new(PropertyId.PagerTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PAGER_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PAGER_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag PagerTelephoneNumberW =
        new(PropertyId.PagerTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PARENT_DISPLAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PARENT_DISPLAY.
    /// </remarks>
    public static readonly PropertyTag ParentDisplayA = new(PropertyId.ParentDisplay, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PARENT_DISPLAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PARENT_DISPLAY.
    /// </remarks>
    public static readonly PropertyTag ParentDisplayW = new(PropertyId.ParentDisplay, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PARENT_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PARENT_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag ParentEntryId = new(PropertyId.ParentEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PARENT_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PARENT_KEY.
    /// </remarks>
    public static readonly PropertyTag ParentKey = new(PropertyId.ParentKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PERSONAL_HOME_PAGE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PERSONAL_HOME_PAGE.
    /// </remarks>
    public static readonly PropertyTag PersonalHomePageA = new(PropertyId.PersonalHomePage, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PERSONAL_HOME_PAGE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PERSONAL_HOME_PAGE.
    /// </remarks>
    public static readonly PropertyTag PersonalHomePageW = new(PropertyId.PersonalHomePage, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY.
    /// </remarks>
    public static readonly PropertyTag PhysicalDeliveryBureauFaxDelivery =
        new(PropertyId.PhysicalDeliveryBureauFaxDelivery, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_PHYSICAL_DELIVERY_MODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PHYSICAL_DELIVERY_MODE.
    /// </remarks>
    public static readonly PropertyTag PhysicalDeliveryMode = new(PropertyId.PhysicalDeliveryMode, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_PHYSICAL_DELIVERY_REPORT_REQUEST.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PHYSICAL_DELIVERY_REPORT_REQUEST.
    /// </remarks>
    public static readonly PropertyTag PhysicalDeliveryReportRequest =
        new(PropertyId.PhysicalDeliveryReportRequest, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag PhysicalForwardingAddress =
        new(PropertyId.PhysicalForwardingAddress, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag PhysicalForwardingAddressRequested =
        new(PropertyId.PhysicalForwardingAddressRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_PHYSICAL_FORWARDING_PROHIBITED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PHYSICAL_FORWARDING_PROHIBITED.
    /// </remarks>
    public static readonly PropertyTag PhysicalForwardingProhibited =
        new(PropertyId.PhysicalForwardingProhibited, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_PHYSICAL_RENDITION_ATTRIBUTES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PHYSICAL_RENDITION_ATTRIBUTES.
    /// </remarks>
    public static readonly PropertyTag PhysicalRenditionAttributes =
        new(PropertyId.PhysicalRenditionAttributes, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_POSTAL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POSTAL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag PostalAddressA = new(PropertyId.PostalAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_POSTAL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POSTAL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag PostalAddressW = new(PropertyId.PostalAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag PostalCodeA = new(PropertyId.PostalCode, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_POSTAL_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POSTAL_CODE.
    /// </remarks>
    public static readonly PropertyTag PostalCodeW = new(PropertyId.PostalCode, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_POST_FOLDER_ENTRIES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_FOLDER_ENTRIES.
    /// </remarks>
    public static readonly PropertyTag PostFolderEntries = new(PropertyId.PostFolderEntries, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_POST_FOLDER_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_FOLDER_NAMES.
    /// </remarks>
    public static readonly PropertyTag PostFolderNamesA = new(PropertyId.PostFolderNames, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_POST_FOLDER_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_FOLDER_NAMES.
    /// </remarks>
    public static readonly PropertyTag PostFolderNamesW = new(PropertyId.PostFolderNames, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_POST_OFFICE_BOX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_OFFICE_BOX.
    /// </remarks>
    public static readonly PropertyTag PostOfficeBoxA = new(PropertyId.PostOfficeBox, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_POST_OFFICE_BOX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_OFFICE_BOX.
    /// </remarks>
    public static readonly PropertyTag PostOfficeBoxW = new(PropertyId.PostOfficeBox, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_POST_REPLY_DENIED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_REPLY_DENIED.
    /// </remarks>
    public static readonly PropertyTag PostReplyDenied = new(PropertyId.PostReplyDenied, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_POST_REPLY_FOLDER_ENTRIES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_REPLY_FOLDER_ENTRIES.
    /// </remarks>
    public static readonly PropertyTag PostReplyFolderEntries =
        new(PropertyId.PostReplyFolderEntries, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_POST_REPLY_FOLDER_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_REPLY_FOLDER_NAMES.
    /// </remarks>
    public static readonly PropertyTag PostReplyFolderNamesA =
        new(PropertyId.PostReplyFolderNames, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_POST_REPLY_FOLDER_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_POST_REPLY_FOLDER_NAMES.
    /// </remarks>
    public static readonly PropertyTag PostReplyFolderNamesW =
        new(PropertyId.PostReplyFolderNames, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PREFERRED_BY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PREFERRED_BY_NAME.
    /// </remarks>
    public static readonly PropertyTag PreferredByNameA = new(PropertyId.PreferredByName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PREFERRED_BY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PREFERRED_BY_NAME.
    /// </remarks>
    public static readonly PropertyTag PreferredByNameW = new(PropertyId.PreferredByName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PREPROCESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PREPROCESS.
    /// </remarks>
    public static readonly PropertyTag Preprocess = new(PropertyId.Preprocess, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_PRIMARY_CAPABILITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PRIMARY_CAPABILITY.
    /// </remarks>
    public static readonly PropertyTag PrimaryCapability = new(PropertyId.PrimaryCapability, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PRIMARY_FAX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PRIMARY_FAX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag PrimaryFaxNumberA = new(PropertyId.PrimaryFaxNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PRIMARY_FAX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PRIMARY_FAX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag PrimaryFaxNumberW = new(PropertyId.PrimaryFaxNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag PrimaryTelephoneNumberA =
        new(PropertyId.PrimaryTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PRIMARY_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag PrimaryTelephoneNumberW =
        new(PropertyId.PrimaryTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PRIORITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PRIORITY.
    /// </remarks>
    public static readonly PropertyTag Priority = new(PropertyId.Priority, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_PROFESSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROFESSION.
    /// </remarks>
    public static readonly PropertyTag ProfessionA = new(PropertyId.Profession, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PROFESSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROFESSION.
    /// </remarks>
    public static readonly PropertyTag ProfessionW = new(PropertyId.Profession, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PROFILE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROFILE_NAME.
    /// </remarks>
    public static readonly PropertyTag ProfileNameA = new(PropertyId.ProfileName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PROFILE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROFILE_NAME.
    /// </remarks>
    public static readonly PropertyTag ProfileNameW = new(PropertyId.ProfileName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PROOF_OF_DELIVERY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROOF_OF_DELIVERY.
    /// </remarks>
    public static readonly PropertyTag ProofOfDelivery = new(PropertyId.ProofOfDelivery, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PROOF_OF_DELIVERY_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROOF_OF_DELIVERY_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag ProofOfDeliveryRequested =
        new(PropertyId.ProofOfDeliveryRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_PROOF_OF_SUBMISSION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROOF_OF_SUBMISSION.
    /// </remarks>
    public static readonly PropertyTag ProofOfSubmission = new(PropertyId.ProofOfSubmission, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PROOF_OF_SUBMISSION_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROOF_OF_SUBMISSION_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag ProofOfSubmissionRequested =
        new(PropertyId.ProofOfSubmissionRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_PROVIDER_DISPLAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROVIDER_DISPLAY.
    /// </remarks>
    public static readonly PropertyTag ProviderDisplayA = new(PropertyId.ProviderDisplay, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PROVIDER_DISPLAY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROVIDER_DISPLAY.
    /// </remarks>
    public static readonly PropertyTag ProviderDisplayW = new(PropertyId.ProviderDisplay, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PROVIDER_DLL_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROVIDER_DLL_NAME.
    /// </remarks>
    public static readonly PropertyTag ProviderDllNameA = new(PropertyId.ProviderDllName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_PROVIDER_DLL_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROVIDER_DLL_NAME.
    /// </remarks>
    public static readonly PropertyTag ProviderDllNameW = new(PropertyId.ProviderDllName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_PROVIDER_ORDINAL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROVIDER_ORDINAL.
    /// </remarks>
    public static readonly PropertyTag ProviderOrdinal = new(PropertyId.ProviderOrdinal, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_PROVIDER_SUBMIT_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROVIDER_SUBMIT_TIME.
    /// </remarks>
    public static readonly PropertyTag ProviderSubmitTime = new(PropertyId.ProviderSubmitTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_PROVIDER_UID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PROVIDER_UID.
    /// </remarks>
    public static readonly PropertyTag ProviderUid = new(PropertyId.ProviderUid, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_PUID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_PUID.
    /// </remarks>
    public static readonly PropertyTag Puid = new(PropertyId.Puid, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_RADIO_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RADIO_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag RadioTelephoneNumberA =
        new(PropertyId.RadioTelephoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RADIO_TELEPHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RADIO_TELEPHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag RadioTelephoneNumberW =
        new(PropertyId.RadioTelephoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingAddrtypeA =
        new(PropertyId.RcvdRepresentingAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingAddrtypeW =
        new(PropertyId.RcvdRepresentingAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingEmailAddressA =
        new(PropertyId.RcvdRepresentingEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingEmailAddressW =
        new(PropertyId.RcvdRepresentingEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingEntryId =
        new(PropertyId.RcvdRepresentingEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_NAME.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingNameA =
        new(PropertyId.RcvdRepresentingName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_NAME.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingNameW =
        new(PropertyId.RcvdRepresentingName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RCVD_REPRESENTING_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RCVD_REPRESENTING_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag RcvdRepresentingSearchKey =
        new(PropertyId.RcvdRepresentingSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_READ_RECEIPT_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_READ_RECEIPT_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag ReadReceiptEntryId = new(PropertyId.ReadReceiptEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_READ_RECEIPT_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_READ_RECEIPT_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag
        ReadReceiptRequested = new(PropertyId.ReadReceiptRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_READ_RECEIPT_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_READ_RECEIPT_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag ReadReceiptSearchKey = new(PropertyId.ReadReceiptSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_RECEIPT_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIPT_TIME.
    /// </remarks>
    public static readonly PropertyTag ReceiptTime = new(PropertyId.ReceiptTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag ReceivedByAddrtypeA = new(PropertyId.ReceivedByAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag ReceivedByAddrtypeW = new(PropertyId.ReceivedByAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag ReceivedByEmailAddressA =
        new(PropertyId.ReceivedByEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag ReceivedByEmailAddressW =
        new(PropertyId.ReceivedByEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag ReceivedByEntryId = new(PropertyId.ReceivedByEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_NAME.
    /// </remarks>
    public static readonly PropertyTag ReceivedByNameA = new(PropertyId.ReceivedByName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_NAME.
    /// </remarks>
    public static readonly PropertyTag ReceivedByNameW = new(PropertyId.ReceivedByName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RECEIVED_BY_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVED_BY_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag ReceivedBySearchKey = new(PropertyId.ReceivedBySearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_RECEIVE_FOLDER_SETTINGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECEIVE_FOLDER_SETTINGS.
    /// </remarks>
    public static readonly PropertyTag ReceiveFolderSettings =
        new(PropertyId.ReceiveFolderSettings, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_CERTIFICATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_CERTIFICATE.
    /// </remarks>
    public static readonly PropertyTag RecipientCertificate = new(PropertyId.RecipientCertificate, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag RecipientDisplayNameA =
        new(PropertyId.RecipientDisplayName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag RecipientDisplayNameW =
        new(PropertyId.RecipientDisplayName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
    /// </remarks>
    public static readonly PropertyTag RecipientNumberForAdviceA =
        new(PropertyId.RecipientNumberForAdvice, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_NUMBER_FOR_ADVICE.
    /// </remarks>
    public static readonly PropertyTag RecipientNumberForAdviceW =
        new(PropertyId.RecipientNumberForAdvice, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_REASSIGNMENT_PROHIBITED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_REASSIGNMENT_PROHIBITED.
    /// </remarks>
    public static readonly PropertyTag RecipientReassignmentProhibited =
        new(PropertyId.RecipientReassignmentProhibited, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_STATUS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_STATUS.
    /// </remarks>
    public static readonly PropertyTag RecipientStatus = new(PropertyId.RecipientStatus, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RECIPIENT_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RECIPIENT_TYPE.
    /// </remarks>
    public static readonly PropertyTag RecipientType = new(PropertyId.RecipientType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_REDIRECTION_HISTORY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REDIRECTION_HISTORY.
    /// </remarks>
    public static readonly PropertyTag RedirectionHistory = new(PropertyId.RedirectionHistory, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REFERRED_BY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REFERRED_BY_NAME.
    /// </remarks>
    public static readonly PropertyTag ReferredByNameA = new(PropertyId.ReferredByName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_REFERRED_BY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REFERRED_BY_NAME.
    /// </remarks>
    public static readonly PropertyTag ReferredByNameW = new(PropertyId.ReferredByName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_REGISTERED_MAIL_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REGISTERED_MAIL_TYPE.
    /// </remarks>
    public static readonly PropertyTag RegisteredMailType = new(PropertyId.RegisteredMailType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RELATED_IPMS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RELATED_IPMS.
    /// </remarks>
    public static readonly PropertyTag RelatedIpms = new(PropertyId.RelatedIpms, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REMOTE_PROGRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REMOTE_PROGRESS.
    /// </remarks>
    public static readonly PropertyTag RemoteProgress = new(PropertyId.RemoteProgress, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_REMOTE_PROGRESS_TEXT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REMOTE_PROGRESS_TEXT.
    /// </remarks>
    public static readonly PropertyTag RemoteProgressTextA = new(PropertyId.RemoteProgressText, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_REMOTE_PROGRESS_TEXT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REMOTE_PROGRESS_TEXT.
    /// </remarks>
    public static readonly PropertyTag RemoteProgressTextW = new(PropertyId.RemoteProgressText, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_REMOTE_VALIDATE_OK.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REMOTE_VALIDATE_OK.
    /// </remarks>
    public static readonly PropertyTag RemoteValidateOk = new(PropertyId.RemoteValidateOk, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_RENDERING_POSITION.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RENDERING_POSITION.
    /// </remarks>
    public static readonly PropertyTag RenderingPosition = new(PropertyId.RenderingPosition, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_REPLY_RECIPIENT_ENTRIES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPLY_RECIPIENT_ENTRIES.
    /// </remarks>
    public static readonly PropertyTag ReplyRecipientEntries =
        new(PropertyId.ReplyRecipientEntries, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REPLY_RECIPIENT_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPLY_RECIPIENT_NAMES.
    /// </remarks>
    public static readonly PropertyTag ReplyRecipientNamesA = new(PropertyId.ReplyRecipientNames, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_REPLY_RECIPIENT_NAMES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPLY_RECIPIENT_NAMES.
    /// </remarks>
    public static readonly PropertyTag ReplyRecipientNamesW = new(PropertyId.ReplyRecipientNames, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_REPLY_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPLY_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag ReplyRequested = new(PropertyId.ReplyRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_REPLY_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPLY_TIME.
    /// </remarks>
    public static readonly PropertyTag ReplyTime = new(PropertyId.ReplyTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_REPORT_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag ReportEntryId = new(PropertyId.ReportEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REPORTING_DL_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORTING_DL_NAME.
    /// </remarks>
    public static readonly PropertyTag ReportingDlName = new(PropertyId.ReportingDlName, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REPORTING_MTA_CERTIFICATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORTING_MTA_CERTIFICATE.
    /// </remarks>
    public static readonly PropertyTag ReportingMtaCertificate =
        new(PropertyId.ReportingMtaCertificate, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REPORT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_NAME.
    /// </remarks>
    public static readonly PropertyTag ReportNameA = new(PropertyId.ReportName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_REPORT_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_NAME.
    /// </remarks>
    public static readonly PropertyTag ReportNameW = new(PropertyId.ReportName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_REPORT_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag ReportSearchKey = new(PropertyId.ReportSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REPORT_TAG.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_TAG.
    /// </remarks>
    public static readonly PropertyTag ReportTag = new(PropertyId.ReportTag, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_REPORT_TEXT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_TEXT.
    /// </remarks>
    public static readonly PropertyTag ReportTextA = new(PropertyId.ReportText, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_REPORT_TEXT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_TEXT.
    /// </remarks>
    public static readonly PropertyTag ReportTextW = new(PropertyId.ReportText, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_REPORT_TIME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REPORT_TIME.
    /// </remarks>
    public static readonly PropertyTag ReportTime = new(PropertyId.ReportTime, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_REQUESTED_DELIVERY_METHOD.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_REQUESTED_DELIVERY_METHOD.
    /// </remarks>
    public static readonly PropertyTag RequestedDeliveryMethod =
        new(PropertyId.RequestedDeliveryMethod, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RESOURCE_FLAGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RESOURCE_FLAGS.
    /// </remarks>
    public static readonly PropertyTag ResourceFlags = new(PropertyId.ResourceFlags, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RESOURCE_METHODS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RESOURCE_METHODS.
    /// </remarks>
    public static readonly PropertyTag ResourceMethods = new(PropertyId.ResourceMethods, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RESOURCE_PATH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RESOURCE_PATH.
    /// </remarks>
    public static readonly PropertyTag ResourcePathA = new(PropertyId.ResourcePath, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RESOURCE_PATH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RESOURCE_PATH.
    /// </remarks>
    public static readonly PropertyTag ResourcePathW = new(PropertyId.ResourcePath, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RESOURCE_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RESOURCE_TYPE.
    /// </remarks>
    public static readonly PropertyTag ResourceType = new(PropertyId.ResourceType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RESPONSE_REQUESTED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RESPONSE_REQUESTED.
    /// </remarks>
    public static readonly PropertyTag ResponseRequested = new(PropertyId.ResponseRequested, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_RESPONSIBILITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RESPONSIBILITY.
    /// </remarks>
    public static readonly PropertyTag Responsibility = new(PropertyId.Responsibility, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_RETURNED_IPM.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RETURNED_IPM.
    /// </remarks>
    public static readonly PropertyTag ReturnedIpm = new(PropertyId.ReturnedIpm, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_ROWID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ROWID.
    /// </remarks>
    public static readonly PropertyTag Rowid = new(PropertyId.Rowid, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_ROW_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_ROW_TYPE.
    /// </remarks>
    public static readonly PropertyTag RowType = new(PropertyId.RowType, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RTF_COMPRESSED.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_COMPRESSED.
    /// </remarks>
    public static readonly PropertyTag RtfCompressed = new(PropertyId.RtfCompressed, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_RTF_IN_SYNC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_IN_SYNC.
    /// </remarks>
    public static readonly PropertyTag RtfInSync = new(PropertyId.RtfInSync, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_RTF_SYNC_BODY_COUNT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_SYNC_BODY_COUNT.
    /// </remarks>
    public static readonly PropertyTag RtfSyncBodyCount = new(PropertyId.RtfSyncBodyCount, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RTF_SYNC_BODY_CRC.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_SYNC_BODY_CRC.
    /// </remarks>
    public static readonly PropertyTag RtfSyncBodyCrc = new(PropertyId.RtfSyncBodyCrc, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RTF_SYNC_BODY_TAG.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_SYNC_BODY_TAG.
    /// </remarks>
    public static readonly PropertyTag RtfSyncBodyTagA = new(PropertyId.RtfSyncBodyTag, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_RTF_SYNC_BODY_TAG.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_SYNC_BODY_TAG.
    /// </remarks>
    public static readonly PropertyTag RtfSyncBodyTagW = new(PropertyId.RtfSyncBodyTag, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_RTF_SYNC_PREFIX_COUNT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_SYNC_PREFIX_COUNT.
    /// </remarks>
    public static readonly PropertyTag RtfSyncPrefixCount = new(PropertyId.RtfSyncPrefixCount, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_RTF_SYNC_TRAILING_COUNT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_RTF_SYNC_TRAILING_COUNT.
    /// </remarks>
    public static readonly PropertyTag RtfSyncTrailingCount = new(PropertyId.RtfSyncTrailingCount, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_SEARCH.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SEARCH.
    /// </remarks>
    public static readonly PropertyTag Search = new(PropertyId.Search, PropertyType.Object);

    /// <summary>
    ///     The MAPI property PR_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag SearchKey = new(PropertyId.SearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SECURITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SECURITY.
    /// </remarks>
    public static readonly PropertyTag Security = new(PropertyId.Security, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_SELECTABLE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SELECTABLE.
    /// </remarks>
    public static readonly PropertyTag Selectable = new(PropertyId.Selectable, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_SENDER_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag SenderAddrtypeA = new(PropertyId.SenderAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SENDER_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag SenderAddrtypeW = new(PropertyId.SenderAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SENDER_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag SenderEmailAddressA = new(PropertyId.SenderEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SENDER_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag SenderEmailAddressW = new(PropertyId.SenderEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SENDER_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag SenderEntryId = new(PropertyId.SenderEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SENDER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_NAME.
    /// </remarks>
    public static readonly PropertyTag SenderNameA = new(PropertyId.SenderName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SENDER_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_NAME.
    /// </remarks>
    public static readonly PropertyTag SenderNameW = new(PropertyId.SenderName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SENDER_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENDER_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag SenderSearchKey = new(PropertyId.SenderSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SEND_INTERNET_ENCODING.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SEND_INTERNET_ENCODING.
    /// </remarks>
    public static readonly PropertyTag SendInternetEncoding = new(PropertyId.SendInternetEncoding, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_SEND_RECALL_REPORT
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SEND_RECALL_REPORT.
    /// </remarks>
    public static readonly PropertyTag SendRecallReport = new(PropertyId.SendRecallReport, PropertyType.Unspecified);

    /// <summary>
    ///     The MAPI property PR_SEND_RICH_INFO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SEND_RICH_INFO.
    /// </remarks>
    public static readonly PropertyTag SendRichInfo = new(PropertyId.SendRichInfo, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_SENSITIVITY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENSITIVITY.
    /// </remarks>
    public static readonly PropertyTag Sensitivity = new(PropertyId.Sensitivity, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_SENTMAIL_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENTMAIL_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag SentmailEntryId = new(PropertyId.SentmailEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingAddrtypeA =
        new(PropertyId.SentRepresentingAddrtype, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_ADDRTYPE.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingAddrtypeW =
        new(PropertyId.SentRepresentingAddrtype, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingEmailAddressA =
        new(PropertyId.SentRepresentingEmailAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_EMAIL_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingEmailAddressW =
        new(PropertyId.SentRepresentingEmailAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingEntryId =
        new(PropertyId.SentRepresentingEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_NAME.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingNameA =
        new(PropertyId.SentRepresentingName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_NAME.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingNameW =
        new(PropertyId.SentRepresentingName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SENT_REPRESENTING_SEARCH_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SENT_REPRESENTING_SEARCH_KEY.
    /// </remarks>
    public static readonly PropertyTag SentRepresentingSearchKey =
        new(PropertyId.SentRepresentingSearchKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SERVICE_DELETE_FILES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_DELETE_FILES.
    /// </remarks>
    public static readonly PropertyTag ServiceDeleteFilesA = new(PropertyId.ServiceDeleteFiles, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SERVICE_DELETE_FILES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_DELETE_FILES.
    /// </remarks>
    public static readonly PropertyTag ServiceDeleteFilesW = new(PropertyId.ServiceDeleteFiles, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SERVICE_DLL_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_DLL_NAME.
    /// </remarks>
    public static readonly PropertyTag ServiceDllNameA = new(PropertyId.ServiceDllName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SERVICE_DLL_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_DLL_NAME.
    /// </remarks>
    public static readonly PropertyTag ServiceDllNameW = new(PropertyId.ServiceDllName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SERVICE_ENTRY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_ENTRY_NAME.
    /// </remarks>
    public static readonly PropertyTag ServiceEntryName = new(PropertyId.ServiceEntryName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SERVICE_EXTRA_UIDS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_EXTRA_UIDS.
    /// </remarks>
    public static readonly PropertyTag ServiceExtraUids = new(PropertyId.ServiceExtraUids, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SERVICE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_NAME.
    /// </remarks>
    public static readonly PropertyTag ServiceNameA = new(PropertyId.ServiceName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SERVICE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_NAME.
    /// </remarks>
    public static readonly PropertyTag ServiceNameW = new(PropertyId.ServiceName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SERVICES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICES.
    /// </remarks>
    public static readonly PropertyTag Services = new(PropertyId.Services, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SERVICE_SUPPORT_FILES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_SUPPORT_FILES.
    /// </remarks>
    public static readonly PropertyTag ServiceSupportFilesA = new(PropertyId.ServiceSupportFiles, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SERVICE_SUPPORT_FILES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_SUPPORT_FILES.
    /// </remarks>
    public static readonly PropertyTag ServiceSupportFilesW = new(PropertyId.ServiceSupportFiles, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SERVICE_UID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SERVICE_UID.
    /// </remarks>
    public static readonly PropertyTag ServiceUid = new(PropertyId.ServiceUid, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SEVEN_BIT_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SEVEN_BIT_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag SevenBitDisplayName = new(PropertyId.SevenBitDisplayName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SMTP_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SMTP_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag SmtpAddressA = new(PropertyId.SmtpAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SMTP_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SMTP_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag SmtpAddressW = new(PropertyId.SmtpAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SPOOLER_STATUS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SPOOLER_STATUS.
    /// </remarks>
    public static readonly PropertyTag SpoolerStatus = new(PropertyId.SpoolerStatus, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_SPOUSE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SPOUSE_NAME.
    /// </remarks>
    public static readonly PropertyTag SpouseNameA = new(PropertyId.SpouseName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SPOUSE_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SPOUSE_NAME.
    /// </remarks>
    public static readonly PropertyTag SpouseNameW = new(PropertyId.SpouseName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_START_DATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_START_DATE.
    /// </remarks>
    public static readonly PropertyTag StartDate = new(PropertyId.StartDate, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_STATE_OR_PROVINCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STATE_OR_PROVINCE.
    /// </remarks>
    public static readonly PropertyTag StateOrProvinceA = new(PropertyId.StateOrProvince, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_STATE_OR_PROVINCE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STATE_OR_PROVINCE.
    /// </remarks>
    public static readonly PropertyTag StateOrProvinceW = new(PropertyId.StateOrProvince, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_STATUS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STATUS.
    /// </remarks>
    public static readonly PropertyTag Status = new(PropertyId.Status, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_STATUS_CODE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STATUS_CODE.
    /// </remarks>
    public static readonly PropertyTag StatusCode = new(PropertyId.StatusCode, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_STATUS_STRING.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STATUS_STRING.
    /// </remarks>
    public static readonly PropertyTag StatusStringA = new(PropertyId.StatusString, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_STATUS_STRING.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STATUS_STRING.
    /// </remarks>
    public static readonly PropertyTag StatusStringW = new(PropertyId.StatusString, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_STORE_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STORE_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag StoreEntryId = new(PropertyId.StoreEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_STORE_PROVIDERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STORE_PROVIDERS.
    /// </remarks>
    public static readonly PropertyTag StoreProviders = new(PropertyId.StoreProviders, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_STORE_RECORD_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STORE_RECORD_KEY.
    /// </remarks>
    public static readonly PropertyTag StoreRecordKey = new(PropertyId.StoreRecordKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_STORE_STATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STORE_STATE.
    /// </remarks>
    public static readonly PropertyTag StoreState = new(PropertyId.StoreState, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_STORE_SUPPORT_MASK.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STORE_SUPPORT_MASK.
    /// </remarks>
    public static readonly PropertyTag StoreSupportMask = new(PropertyId.StoreSupportMask, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_STREET_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STREET_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag StreetAddressA = new(PropertyId.StreetAddress, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_STREET_ADDRESS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_STREET_ADDRESS.
    /// </remarks>
    public static readonly PropertyTag StreetAddressW = new(PropertyId.StreetAddress, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SUBFOLDERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUBFOLDERS.
    /// </remarks>
    public static readonly PropertyTag Subfolders = new(PropertyId.Subfolders, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_SUBJECT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUBJECT.
    /// </remarks>
    public static readonly PropertyTag SubjectA = new(PropertyId.Subject, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SUBJECT.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUBJECT.
    /// </remarks>
    public static readonly PropertyTag SubjectW = new(PropertyId.Subject, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SUBJECT_IPM.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUBJECT_IPM.
    /// </remarks>
    public static readonly PropertyTag SubjectIpm = new(PropertyId.SubjectIpm, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_SUBJECT_PREFIX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUBJECT_PREFIX.
    /// </remarks>
    public static readonly PropertyTag SubjectPrefixA = new(PropertyId.SubjectPrefix, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SUBJECT_PREFIX.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUBJECT_PREFIX.
    /// </remarks>
    public static readonly PropertyTag SubjectPrefixW = new(PropertyId.SubjectPrefix, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SUBMIT_FLAGS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUBMIT_FLAGS.
    /// </remarks>
    public static readonly PropertyTag SubmitFlags = new(PropertyId.SubmitFlags, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_SUPERSEDES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUPERSEDES.
    /// </remarks>
    public static readonly PropertyTag SupersedesA = new(PropertyId.Supersedes, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SUPERSEDES.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUPERSEDES.
    /// </remarks>
    public static readonly PropertyTag SupersedesW = new(PropertyId.Supersedes, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SUPPLEMENTARY_INFO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUPPLEMENTARY_INFO.
    /// </remarks>
    public static readonly PropertyTag SupplementaryInfoA = new(PropertyId.SupplementaryInfo, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SUPPLEMENTARY_INFO.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SUPPLEMENTARY_INFO.
    /// </remarks>
    public static readonly PropertyTag SupplementaryInfoW = new(PropertyId.SupplementaryInfo, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_SURNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SURNAME.
    /// </remarks>
    public static readonly PropertyTag SurnameA = new(PropertyId.Surname, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_SURNAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_SURNAME.
    /// </remarks>
    public static readonly PropertyTag SurnameW = new(PropertyId.Surname, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_TELEX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TELEX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag TelexNumberA = new(PropertyId.TelexNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_TELEX_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TELEX_NUMBER.
    /// </remarks>
    public static readonly PropertyTag TelexNumberW = new(PropertyId.TelexNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_TEMPLATEID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TEMPLATEID.
    /// </remarks>
    public static readonly PropertyTag Templateid = new(PropertyId.Templateid, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_TITLE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TITLE.
    /// </remarks>
    public static readonly PropertyTag TitleA = new(PropertyId.Title, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_TITLE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TITLE.
    /// </remarks>
    public static readonly PropertyTag TitleW = new(PropertyId.Title, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_TNEF_CORRELATION_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TNEF_CORRELATION_KEY.
    /// </remarks>
    public static readonly PropertyTag TnefCorrelationKey = new(PropertyId.TnefCorrelationKey, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag TransmitableDisplayNameA =
        new(PropertyId.TransmitableDisplayName, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TRANSMITABLE_DISPLAY_NAME.
    /// </remarks>
    public static readonly PropertyTag TransmitableDisplayNameW =
        new(PropertyId.TransmitableDisplayName, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_TRANSPORT_KEY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TRANSPORT_KEY.
    /// </remarks>
    public static readonly PropertyTag TransportKey = new(PropertyId.TransportKey, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
    /// </remarks>
    public static readonly PropertyTag TransportMessageHeadersA =
        new(PropertyId.TransportMessageHeaders, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TRANSPORT_MESSAGE_HEADERS.
    /// </remarks>
    public static readonly PropertyTag TransportMessageHeadersW =
        new(PropertyId.TransportMessageHeaders, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_TRANSPORT_PROVIDERS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TRANSPORT_PROVIDERS.
    /// </remarks>
    public static readonly PropertyTag TransportProviders = new(PropertyId.TransportProviders, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_TRANSPORT_STATUS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TRANSPORT_STATUS.
    /// </remarks>
    public static readonly PropertyTag TransportStatus = new(PropertyId.TransportStatus, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_TTYDD_PHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TTYDD_PHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag TtytddPhoneNumberA = new(PropertyId.TtytddPhoneNumber, PropertyType.String8);

    /// <summary>
    ///     The MAPI property PR_TTYDD_PHONE_NUMBER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TTYDD_PHONE_NUMBER.
    /// </remarks>
    public static readonly PropertyTag TtytddPhoneNumberW = new(PropertyId.TtytddPhoneNumber, PropertyType.Unicode);

    /// <summary>
    ///     The MAPI property PR_TYPE_OF_MTS_USER.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_TYPE_OF_MTS_USER.
    /// </remarks>
    public static readonly PropertyTag TypeOfMtsUser = new(PropertyId.TypeOfMtsUser, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_USER_CERTIFICATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_USER_CERTIFICATE.
    /// </remarks>
    public static readonly PropertyTag UserCertificate = new(PropertyId.UserCertificate, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_USER_X509_CERTIFICATE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_USER_X509_CERTIFICATE.
    /// </remarks>
    public static readonly PropertyTag UserX509Certificate = new(PropertyId.UserX509Certificate, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_VALID_FOLDER_MASK.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_VALID_FOLDER_MASK.
    /// </remarks>
    public static readonly PropertyTag ValidFolderMask = new(PropertyId.ValidFolderMask, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_VIEWS_ENTRYID.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_VIEWS_ENTRYID.
    /// </remarks>
    public static readonly PropertyTag ViewsEntryId = new(PropertyId.ViewsEntryId, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_WEDDING_ANNIVERSARY.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_WEDDING_ANNIVERSARY.
    /// </remarks>
    public static readonly PropertyTag WeddingAnniversary = new(PropertyId.WeddingAnniversary, PropertyType.SysTime);

    /// <summary>
    ///     The MAPI property PR_X400_CONTENT_TYPE.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_X400_CONTENT_TYPE.
    /// </remarks>
    public static readonly PropertyTag X400ContentType = new(PropertyId.X400ContentType, PropertyType.Binary);

    /// <summary>
    ///     The MAPI property PR_X400_DEFERRED_DELIVERY_CANCEL.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_X400_DEFERRED_DELIVERY_CANCEL.
    /// </remarks>
    public static readonly PropertyTag X400DeferredDeliveryCancel =
        new(PropertyId.X400DeferredDeliveryCancel, PropertyType.Boolean);

    /// <summary>
    ///     The MAPI property PR_XPOS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_XPOS.
    /// </remarks>
    public static readonly PropertyTag Xpos = new(PropertyId.Xpos, PropertyType.Long);

    /// <summary>
    ///     The MAPI property PR_YPOS.
    /// </summary>
    /// <remarks>
    ///     The MAPI property PR_YPOS.
    /// </remarks>
    public static readonly PropertyTag Ypos = new(PropertyId.Ypos, PropertyType.Long);
    #endregion

    #region Consts
    private const PropertyId NamedMin = unchecked((PropertyId)0x8000);
    private const PropertyId NamedMax = unchecked((PropertyId)0xFFFE);
    private const short MultiValuedFlag = (short)PropertyType.MultiValued;
    #endregion

    #region Properties
    /// <summary>
    ///     Get the property identifier.
    /// </summary>
    /// <remarks>
    ///     Gets the property identifier.
    /// </remarks>
    /// <value>The identifier.</value>
    internal PropertyId Id { get; }

    /// <summary>
    ///     Get a value indicating whether or not the property contains multiple values.
    /// </summary>
    /// <remarks>
    ///     Gets a value indicating whether or not the property contains multiple values.
    /// </remarks>
    /// <value><c>true</c> if the property contains multiple values; otherwise, <c>false</c>.</value>
    public bool IsMultiValued => ((short)TnefType & MultiValuedFlag) != 0;

    /// <summary>
    ///     Get a value indicating whether or not the property has a special name.
    /// </summary>
    /// <remarks>
    ///     Gets a value indicating whether or not the property has a special name.
    /// </remarks>
    /// <value><c>true</c> if the property has a special name; otherwise, <c>false</c>.</value>
    public bool IsNamed => (int)Id >= (int)NamedMin && (int)Id <= (int)NamedMax;

    /// <summary>
    ///     Get a value indicating whether the property value type is valid.
    /// </summary>
    /// <remarks>
    ///     Gets a value indicating whether the property value type is valid.
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
    ///     Get the property's value type (including the multi-valued bit).
    /// </summary>
    /// <remarks>
    ///     Gets the property's value type (including the multi-valued bit).
    /// </remarks>
    /// <value>The property's value type.</value>
    public PropertyType TnefType { get; }

    /// <summary>
    ///     Get the type of the value that the property contains.
    /// </summary>
    /// <remarks>
    ///     Gets the type of the value that the property contains.
    /// </remarks>
    /// <value>The type of the value.</value>
    public PropertyType ValueTnefType => (PropertyType)((short)TnefType & ~MultiValuedFlag);
    #endregion

    #region Constructors
    /// <summary>
    ///     Initialize a new instance of the <see cref="PropertyTag" /> struct.
    /// </summary>
    /// <remarks>
    ///     Creates a new <see cref="PropertyTag" /> based on a 32-bit integer tag as read from
    ///     a TNEF stream.
    /// </remarks>
    /// <param name="tag">The property tag.</param>
    internal PropertyTag(int tag)
    {
        TnefType = (PropertyType)((tag >> 16) & 0xFFFF);
        Id = (PropertyId)(tag & 0xFFFF);
    }

    internal PropertyTag(PropertyId id, PropertyType type, bool multiValue)
    {
        TnefType = (PropertyType)((ushort)type | (multiValue ? MultiValuedFlag : 0));
        Id = id;
    }

    /// <summary>
    ///     Initialize a new instance of the <see cref="PropertyTag" /> struct.
    /// </summary>
    /// <remarks>
    ///     Creates a new <see cref="PropertyTag" /> based on a <see cref="PropertyId" />
    ///     and <see cref="PropertyType" />.
    /// </remarks>
    /// <param name="id">The property identifier.</param>
    /// <param name="type">The property type.</param>
    internal PropertyTag(PropertyId id, PropertyType type)
    {
        TnefType = type;
        Id = id;
    }
    #endregion

    #region Operators
    /// <summary>
    ///     Cast an integer tag value into a TNEF property tag.
    /// </summary>
    /// <remarks>
    ///     Casts an integer tag value into a TNEF property tag.
    /// </remarks>
    /// <returns>A <see cref="PropertyTag" /> that represents the integer tag value.</returns>
    /// <param name="tag">The integer tag value.</param>
    public static implicit operator PropertyTag(int tag)
    {
        return new PropertyTag(tag);
    }

    /// <summary>
    ///     Cast a TNEF property tag into a 32-bit integer value.
    /// </summary>
    /// <remarks>
    ///     Casts a TNEF property tag into a 32-bit integer value.
    /// </remarks>
    /// <returns>A 32-bit integer value representing the TNEF property tag.</returns>
    /// <param name="tag">The TNEF property tag.</param>
    public static implicit operator int(PropertyTag tag)
    {
        return ((ushort)tag.TnefType << 16) | (ushort)tag.Id;
    }
    #endregion

    #region GetHashCode
    /// <summary>
    ///     Serves as a hash function for a <see cref="PropertyTag" /> object.
    /// </summary>
    /// <remarks>
    ///     Serves as a hash function for a <see cref="PropertyTag" /> object.
    /// </remarks>
    /// <returns>
    ///     A hash code for this instance that is suitable for use in hashing algorithms
    ///     and data structures such as a hash table.
    /// </returns>
    public override int GetHashCode()
    {
        return ((int)this).GetHashCode();
    }
    #endregion

    #region Equals
    /// <summary>
    ///     Determine whether the specified <see cref="object" /> is equal to the current <see cref="PropertyTag" />.
    /// </summary>
    /// <remarks>
    ///     Determines whether the specified <see cref="object" /> is equal to the current <see cref="PropertyTag" />.
    /// </remarks>
    /// <param name="obj">The <see cref="object" /> to compare with the current <see cref="PropertyTag" />.</param>
    /// <returns>
    ///     <c>true</c> if the specified <see cref="object" /> is equal to the current
    ///     <see cref="PropertyTag" />; otherwise, <c>false</c>.
    /// </returns>
    public override bool Equals(object obj)
    {
        return obj is PropertyTag tag
               && tag.Id == Id
               && tag.TnefType == TnefType;
    }
    #endregion

    #region ToString
    /// <summary>
    ///     Return a <see cref="string" /> that represents the current <see cref="PropertyTag" />.
    /// </summary>
    /// <remarks>
    ///     Returns a <see cref="string" /> that represents the current <see cref="PropertyTag" />.
    /// </remarks>
    /// <returns>A <see cref="string" /> that represents the current <see cref="PropertyTag" />.</returns>
    public override string ToString()
    {
        return $"{Id} ({ValueTnefType})";
    }
    #endregion

    #region ToUnicode
    /// <summary>
    ///     Return a new <see cref="PropertyTag" /> where the type has been changed to <see cref="PropertyType.Unicode" />.
    /// </summary>
    /// <remarks>
    ///     Returns a new <see cref="PropertyTag" /> where the type has been changed to <see cref="PropertyType.Unicode" />.
    /// </remarks>
    /// <returns>The unicode equivalent of the property tag.</returns>
    public PropertyTag ToUnicode()
    {
        var unicode = (PropertyType)(((short)TnefType & MultiValuedFlag) | (short)PropertyType.Unicode);

        return new PropertyTag(Id, unicode);
    }
    #endregion
}