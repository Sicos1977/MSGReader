﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using MsgReader.Helpers;
using MsgReader.Mime.Decode;

namespace MsgReader.Mime.Header;

/// <summary>
///     Class that can parse different fields in the header sections of a MIME message.
/// </summary>
internal static class HeaderFieldParser
{
    #region ParseContentTransferEncoding
    /// <summary>
    ///     Parses the Content-Transfer-Encoding header.
    /// </summary>
    /// <param name="headerValue">The value for the header to be parsed</param>
    /// <returns>A <see cref="ContentTransferEncoding" /></returns>
    /// <exception cref="ArgumentNullException">If <paramref name="headerValue" /> is <see langword="null" /></exception>
    /// <exception cref="ArgumentException">
    ///     If the <paramref name="headerValue" /> could not be parsed to a
    ///     <see cref="ContentTransferEncoding" />
    /// </exception>
    public static ContentTransferEncoding ParseContentTransferEncoding(string headerValue)
    {
        if (headerValue == null)
            throw new ArgumentNullException(nameof(headerValue));

        switch (headerValue.Trim().ToUpperInvariant())
        {
            case "7BIT":
                return ContentTransferEncoding.SevenBit;

            case "8BIT":
                return ContentTransferEncoding.EightBit;

            case "QUOTED-PRINTABLE":
                return ContentTransferEncoding.QuotedPrintable;

            case "BASE64":
                return ContentTransferEncoding.Base64;

            case "BINARY":
                return ContentTransferEncoding.Binary;

            case "UUENCODE":
                return ContentTransferEncoding.UUEncode;

            // If a wrong argument is passed to this parser method, then we assume
            // default encoding, which is SevenBit.
            // This is to ensure that we do not throw exceptions, even if the email not MIME valid.
            default:
                return ContentTransferEncoding.SevenBit;
        }
    }
    #endregion

    #region ParseImportance
    /// <summary>
    ///     Parses an ImportanceType from a given Importance header value.
    /// </summary>
    /// <param name="headerValue">The value to be parsed</param>
    /// <returns>A <see cref="MailPriority" />. If the <paramref name="headerValue" /> is not recognized, Normal is returned.</returns>
    /// <exception cref="ArgumentNullException">If <paramref name="headerValue" /> is <see langword="null" /></exception>
    public static MailPriority ParseImportance(string headerValue)
    {
        if (headerValue == null)
            throw new ArgumentNullException(nameof(headerValue));

        switch (headerValue.ToUpperInvariant())
        {
            case "5":
            case "HIGH":
                return MailPriority.High;

            case "3":
            case "NORMAL":
                return MailPriority.Normal;

            case "1":
            case "LOW":
                return MailPriority.Low;

            default:
                return MailPriority.Normal;
        }
    }
    #endregion

    #region ParseContentType
    /// <summary>
    ///     Parses a the value for the header Content-Type to
    ///     a <see cref="ContentType" /> object.
    /// </summary>
    /// <param name="headerValue">The value to be parsed</param>
    /// <returns>A <see cref="ContentType" /> object</returns>
    /// <exception cref="ArgumentNullException">If <paramref name="headerValue" /> is <see langword="null" /></exception>
    public static ContentType ParseContentType(string headerValue)
    {
        if (headerValue == null)
            throw new ArgumentNullException(nameof(headerValue));

        // We create an empty Content-Type which we will fill in when we see the values
        var contentType = new ContentType();

        // Now decode the parameters
        var parameters = Rfc2231Decoder.Decode(headerValue);

        var isMediaTypeProcessed = false;
        foreach (var keyValuePair in parameters)
        {
            var key = keyValuePair.Key.ToUpperInvariant().Trim();
            var value = Utility.RemoveQuotesIfAny(keyValuePair.Value.Trim());
            value = value.TrimStart('/');
            value = value.TrimStart('\\');

            switch (key)
            {
                case "":
                    // This is the MediaType - it has no key since it is the first one mentioned in the
                    // headerValue and has no = in it.

                    if (!isMediaTypeProcessed)
                    {
                        // Check for illegal content-type 
                        var v = value.ToUpperInvariant().Trim('\0');
                        if (v.Equals("TEXT") || v.Equals("TEXT/"))
                            value = "text/plain";

                        try
                        {
                            contentType.MediaType = value;
                        }
                        catch
                        {
                            Logger.WriteToLog(
                                $"The mediatype '{value}' is invalid, using application/octet-stream instead");
                            contentType.MediaType = "application/octet-stream";
                        }

                        isMediaTypeProcessed = true;
                    }

                    break;

                case "BOUNDARY":
                    contentType.Boundary = value;
                    break;

                case "CHARSET":
                    contentType.CharSet = value;
                    break;

                case "NAME":
                    contentType.Name = EncodedWord.Decode(value);
                    break;

                default:
                    // This is to shut up the code help that is saying that contentType.Parameters
                    // can be null - which it cant!
                    if (contentType.Parameters == null)
                        throw new Exception("The ContentType parameters property is null. This will never be thrown.");

                    // We add the unknown value to our parameters list
                    // "Known" unknown values are:
                    // - title
                    // - report-type
                    contentType.Parameters.Add(key, value);
                    break;
            }
        }

        return contentType;
    }
    #endregion

    #region ParseContentDisposition
    /// <summary>
    ///     Parses the value for the header Content-Disposition to a <see cref="ContentDisposition" /> object.
    /// </summary>
    /// <param name="headerValue">The value to be parsed</param>
    /// <returns>A <see cref="ContentDisposition" /> object</returns>
    /// <exception cref="ArgumentNullException">If <paramref name="headerValue" /> is <see langword="null" /></exception>
    public static ContentDisposition ParseContentDisposition(string headerValue)
    {
        if (headerValue == null)
            throw new ArgumentNullException(nameof(headerValue));

        // See http://www.ietf.org/rfc/rfc2183.txt for RFC definition

        // Create empty ContentDisposition - we will fill in details as we read them
        var contentDisposition = new ContentDisposition();

        // Now decode the parameters
        var parameters = Rfc2231Decoder.Decode(headerValue);

        foreach (var keyValuePair in parameters)
        {
            var key = keyValuePair.Key.ToUpperInvariant().Trim();
            var value = Utility.RemoveQuotesIfAny(keyValuePair.Value.Trim());
            switch (key)
            {
                case "":
                    // This is the Disposition type - it has no key since it is the first one
                    // and has no = in it.
                    contentDisposition.DispositionType = value;
                    break;

                // The correct name of the parameter is filename, but some emails also contains the parameter
                // name, which also holds the name of the file. Therefore, we use both names for the same field.
                case "NAME":
                case "FILENAME":
				case "REMOTE-IMAGE":
                    // The filename might be in quotes, and it might be encoded-word encoded
                    contentDisposition.FileName = EncodedWord.Decode(value);
                    break;

                case "CREATION-DATE":
                    var creationDate = Rfc2822DateTime.StringToDate(value);
                    contentDisposition.CreationDate = creationDate;
                    break;

                case "MODIFICATION-DATE":
                case "MODIFICATION-DATE-PARM":
                    var modificationDate = Rfc2822DateTime.StringToDate(value);
                    contentDisposition.ModificationDate = modificationDate;
                    break;

                case "READ-DATE":
                    var readDate = Rfc2822DateTime.StringToDate(value);
                    contentDisposition.ReadDate = readDate;
                    break;

                case "SIZE":
                    contentDisposition.Size = SizeParser.Parse(value);
                    break;

                case "ALT":
                    contentDisposition.Parameters.Add(key, value);
                    break;

                case "CHARSET": // ignoring invalid parameter in Content-Disposition
                case "VOICE":
                    break;

                default:
                    if (!key.StartsWith("X-"))
                        throw new ArgumentException($"Unknown parameter in Content-Disposition. Ask developer to fix! Parameter: {key}");
                    contentDisposition.Parameters.Add(key, value);
                    break;
            }
        }

        return contentDisposition;
    }
    #endregion

    #region ParseId
    /// <summary>
    ///     Parses an ID like Message-Id and Content-Id.<br />
    ///     Example:<br />
    ///     <c>&lt;test@test.com&gt;</c><br />
    ///     into<br />
    ///     <c>test@test.com</c>
    /// </summary>
    /// <param name="headerValue">The id to parse</param>
    /// <returns>A parsed ID</returns>
    public static string ParseId(string headerValue)
    {
        // Remove whitespace in front and behind since
        // whitespace is allowed there
        // Remove the last > and the first <
        return headerValue.Trim().TrimEnd('>').TrimStart('<');
    }
    #endregion

    #region ParseMultipleIDs
    /// <summary>
    ///     Parses multiple IDs from a single string like In-Reply-To.
    /// </summary>
    /// <param name="headerValue">The value to parse</param>
    /// <returns>A list of IDs</returns>
    public static List<string> ParseMultipleIDs(string headerValue)
    {
        // Split the string by >
        // We cannot use ' ' (space) here since this is a possible value:
        // <test@test.com><test2@test.com>
        var ids = headerValue.Trim().Split(['>'], StringSplitOptions.RemoveEmptyEntries);
        return ids.Select(ParseId).ToList();
    }
    #endregion
}