using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using MsgReader.Mime.Decode;

namespace MsgReader.Mime.Header
{
    /// <summary>
    /// This class is used for RFC compliant email addresses.<br/>
    /// <br/>
    /// The class cannot be instantiated from outside the library.
    /// </summary>
    /// <remarks>
    /// The <seealso cref="MailAddress"/> does not cover all the possible formats 
    /// for <a href="http://tools.ietf.org/html/rfc5322#section-3.4">RFC 5322 section 3.4</a> compliant email addresses.
    /// This class is used as an address wrapper to account for that deficiency.
    /// </remarks>
    public class RfcMailAddress
    {
        #region Properties
        ///<summary>
        /// The email address of this <see cref="RfcMailAddress"/><br/>
        /// It is possibly string.Empty since RFC mail addresses does not require an email address specified.
        ///</summary>
        ///<example>
        /// Example header with email address:<br/>
        /// To: <c>Test test@mail.com</c><br/>
        /// Address will be <c>test@mail.com</c><br/>
        ///</example>
        ///<example>
        /// Example header without email address:<br/>
        /// To: <c>Test</c><br/>
        /// Address will be <see cref="string.Empty"/>.
        ///</example>
        public string Address { get; private set; }

        ///<summary>
        /// The display name of this <see cref="RfcMailAddress"/><br/>
        /// It is possibly <see cref="string.Empty"/> since RFC mail addresses does not require a display name to be specified.
        ///</summary>
        ///<example>
        /// Example header with display name:<br/>
        /// To: <c>Test test@mail.com</c><br/>
        /// DisplayName will be <c>Test</c>
        ///</example>
        ///<example>
        /// Example header without display name:<br/>
        /// To: <c>test@test.com</c><br/>
        /// DisplayName will be <see cref="string.Empty"/>
        ///</example>
        public string DisplayName { get; private set; }

        /// <summary>
        /// This is the Raw string used to describe the <see cref="RfcMailAddress"/>.
        /// </summary>
        public string Raw { get; private set; }

        /// <summary>
        /// The <see cref="MailAddress"/> associated with the <see cref="RfcMailAddress"/>. 
        /// </summary>
        /// <remarks>
        /// The value of this property can be <see lanword="null"/> in instances where the <see cref="MailAddress"/> cannot represent the address properly.<br/>
        /// Use <see cref="HasValidMailAddress"/> property to see if this property is valid.
        /// </remarks>
        public MailAddress MailAddress { get; private set; }

        /// <summary>
        /// Specifies if the object contains a valid <see cref="MailAddress"/> reference.
        /// </summary>
        public bool HasValidMailAddress
        {
            get { return MailAddress != null; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Constructs an <see cref="RfcMailAddress"/> object from a <see cref="MailAddress"/> object.<br/>
        /// This constructor is used when we were able to construct a <see cref="MailAddress"/> from a string.
        /// </summary>
        /// <param name="mailAddress">The address that <paramref name="raw"/> was parsed into</param>
        /// <param name="raw">The raw unparsed input which was parsed into the <paramref name="mailAddress"/></param>
        /// <exception cref="ArgumentNullException">If <paramref name="mailAddress"/> or <paramref name="raw"/> is <see langword="null"/></exception>
        private RfcMailAddress(MailAddress mailAddress, string raw)
        {
            if (mailAddress == null)
                throw new ArgumentNullException("mailAddress");

            if (raw == null)
                throw new ArgumentNullException("raw");

            MailAddress = mailAddress;
            Address = mailAddress.Address;
            DisplayName = mailAddress.DisplayName;
            Raw = raw;
        }

        /// <summary>
        /// When we were unable to parse a string into a <see cref="MailAddress"/>, this constructor can be
        /// used. The Raw string is then used as the <see cref="DisplayName"/>.
        /// </summary>
        /// <param name="raw">The raw unparsed input which could not be parsed</param>
        /// <exception cref="ArgumentNullException">If <paramref name="raw"/> is <see langword="null"/></exception>
        private RfcMailAddress(string raw)
        {
            if (raw == null)
                throw new ArgumentNullException("raw");

            MailAddress = null;
            Address = string.Empty;
            DisplayName = raw;
            Raw = raw;
        }
        #endregion

        #region ToString
        /// <summary>
        /// A string representation of the <see cref="RfcMailAddress"/> object
        /// </summary>
        /// <returns>Returns the string representation for the object</returns>
        public override string ToString()
        {
            return HasValidMailAddress ? MailAddress.ToString() : Raw;
        }
        #endregion

        #region Parsing
        /// <summary>
        /// Parses an email address from a MIME header<br/>
        /// <br/>
        /// Examples of input:
        /// <c>Eksperten mailrobot &lt;noreply@mail.eksperten.dk&gt;</c><br/>
        /// <c>"Eksperten mailrobot" &lt;noreply@mail.eksperten.dk&gt;</c><br/>
        /// <c>&lt;noreply@mail.eksperten.dk&gt;</c><br/>
        /// <c>noreply@mail.eksperten.dk</c><br/>
        /// <br/>
        /// It might also contain encoded text, which will then be decoded.
        /// </summary>
        /// <param name="input">The value to parse out and email and/or a username</param>
        /// <returns>A <see cref="RfcMailAddress"/></returns>
        /// <exception cref="ArgumentNullException">If <paramref name="input"/> is <see langword="null"/></exception>
        /// <remarks>
        /// <see href="http://tools.ietf.org/html/rfc5322#section-3.4">RFC 5322 section 3.4</see> for more details on email syntax.<br/>
        /// <see cref="EncodedWord.Decode">For more information about encoded text</see>.
        /// </remarks>
        internal static RfcMailAddress ParseMailAddress(string input)
        {
            if (input == null)
                throw new ArgumentNullException("input");

            // Decode the value, if it was encoded
            input = EncodedWord.Decode(input.Trim());

            //Remove any redundant sets of angle brackets around the email address
            var lastOpenAngleBracketIdx = input.LastIndexOf('<');
            var lastCloseAngleBracketIdx = input.LastIndexOf('>');

            //Find the index of the first angle bracket in this series of angle brackets, e.g "a>b" <<blah@email.com>> wouldn't find the angle bracket in the display name
            var firstOpenAngleBracketIdx = lastOpenAngleBracketIdx;
            var firstCloseAngleBracketIdx = lastCloseAngleBracketIdx;

            while (firstOpenAngleBracketIdx > 0 && //There is a character before the last open angle bracket
                input[firstOpenAngleBracketIdx - 1] == '<' && //The character before the last open angle bracket is another open angle bracket
                input[firstCloseAngleBracketIdx - 1] == '>') //The character before the last close angle bracket is another close angle bracket
            {
                //Update the first angle bracket indices
                firstOpenAngleBracketIdx--;
                firstCloseAngleBracketIdx--;
            }

            //If the email address in the input string is enclosed in multiple angle brackets
            if (firstOpenAngleBracketIdx != lastOpenAngleBracketIdx)
            {
                //Remove the multiple angle brackets surrounding the email address from the input string leaving just a single set
                input = input.Substring(0, firstOpenAngleBracketIdx) + //Part before any angle brackets (display name if there is one)
                    input.Substring(lastOpenAngleBracketIdx, firstCloseAngleBracketIdx - lastOpenAngleBracketIdx + 1); //actual email address, including one angle bracket either side
            }

            // Find the location of the email address
            var indexStartEmail = input.LastIndexOf('<');
            var indexEndEmail = input.LastIndexOf('>');

            try
            {
                if (indexStartEmail >= 0 && indexEndEmail >= 0)
                {
                    // Check if there is a username in front of the email address
                    var username = indexStartEmail > 0 ? input.Substring(0, indexStartEmail).Trim() : string.Empty;

                    // Parse out the email address without the "<"  and ">"
                    indexStartEmail = indexStartEmail + 1;
                    var emailLength = indexEndEmail - indexStartEmail;
                    var emailAddress = input.Substring(indexStartEmail, emailLength).Trim();

                    // There has been cases where there was no emailaddress between the < and >
                    if (!string.IsNullOrEmpty(emailAddress))
                    {
                        // If the username is quoted, MailAddress' constructor will remove them for us
                        return new RfcMailAddress(new MailAddress(emailAddress, username), input);
                    }
                }

                // This might be on the form noreply@mail.eksperten.dk
                // Check if there is an email, if notm there is no need to try
                if (input.Contains("@"))
                    return new RfcMailAddress(new MailAddress(input), input);
            }
            catch (FormatException)
            {
                // Sometimes invalid emails are sent, like sqlmap-user@sourceforge.net. (last period is illigal)
            }

            // It could be that the format used was simply a name
            // which is indeed valid according to the RFC
            // Example:
            // Eksperten mailrobot
            return new RfcMailAddress(input);
        }

        /// <summary>
        /// Parses input of the form<br/>
        /// <c>Eksperten mailrobot &lt;noreply@mail.eksperten.dk&gt;, ...</c><br/>
        /// to a list of RFCMailAddresses
        /// </summary>
        /// <param name="input">The input that is a comma-separated list of EmailAddresses to parse</param>
        /// <returns>A List of <seealso cref="RfcMailAddress"/> objects extracted from the <paramref name="input"/> parameter.</returns>
        /// <exception cref="ArgumentNullException">If <paramref name="input"/> is <see langword="null"/></exception>
        internal static List<RfcMailAddress> ParseMailAddresses(string input)
        {
            if (input == null)
                throw new ArgumentNullException("input");

            // MailAddresses are split by commas
            IEnumerable<string> mailAddresses = Utility.SplitStringWithCharNotInsideQuotes(input, ',');

            // Parse each of these
            return mailAddresses.Select(ParseMailAddress).ToList();
        }
        #endregion
    }
}