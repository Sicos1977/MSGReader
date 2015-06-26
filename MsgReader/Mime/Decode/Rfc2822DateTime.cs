using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MsgReader.Mime.Decode
{
    /// <summary>
    /// Class used to decode RFC 2822 Date header fields.
    /// </summary>
    internal static class Rfc2822DateTime
    {
        #region Consts
        //Constants
        /// <summary>
        /// Timezone formats that aren't +-hhmm, e.g. UTC, or K. See MatchEvaluator method for conversions
        /// </summary>
        private const string RegexOldTimezoneFormats = @"\b((UT|GMT|EST|EDT|CST|CDT|MST|MDT|PST|MSK|PDT)|([A-IK-Za-ik-z]))\b";

        /// <summary>
        /// Matches any +=hhmm timezone offset, e.g. +0100
        /// </summary>
        private const string RegexNewTimezoneFormats = @"[\+-](?<hours>\d\d)(?<minutes>\d\d)";
        #endregion

        #region Fields
        /// <summary>
        /// Custom DateTime formats - will be tried if cannot parse the dateInput string using the default method
        ///	 Specified using formats at http://msdn.microsoft.com/en-us/library/8kb3ddd4%28v=vs.110%29.aspx
        ///	 One format per string in the array
        /// </summary>
        public static string[] CustomDateTimeFormats { private get; set; }
        #endregion

        #region StringToDate
        /// <summary>
        /// Converts a string in RFC 2822 format into a <see cref="DateTime"/> object
        /// </summary>
        /// <param name="inputDate">The date to convert</param>
        /// <returns>
        /// A valid <see cref="DateTime"/> object, which represents the same time as the string that was converted. 
        /// If <paramref name="inputDate"/> is not a valid date representation, then <see cref="DateTime.MinValue"/> is returned.
        /// </returns>
        /// <exception cref="ArgumentNullException">If <paramref name="inputDate"/> is <see langword="null"/></exception>
        /// <exception cref="ArgumentException">If the <paramref name="inputDate"/> could not be parsed into a <see cref="DateTime"/> object</exception>
        public static DateTime StringToDate(string inputDate)
        {
            if (inputDate == null)
                throw new ArgumentNullException("inputDate");

            // Handle very wrong date time format: Tue Feb 18 10:23:30 2014 (MSK)
            inputDate = FixSpecialCases(inputDate);

            // Old date specification allows comments and a lot of whitespace
            inputDate = StripCommentsAndExcessWhitespace(inputDate);

            try
            {
                // Extract the DateTime
                var dateTime = ExtractDateTime(inputDate);

                // Bail if we could not parse the date
                if (dateTime == DateTime.MinValue)
                    return dateTime;

                // Convert the date into UTC
                dateTime = new DateTime(dateTime.Ticks, DateTimeKind.Utc);

                // Adjust according to the time zone
                dateTime = AdjustTimezone(dateTime, inputDate);

                // Return the parsed date
                return dateTime;
            }
            catch (FormatException e) // Convert.ToDateTime() Failure
            {
                throw new ArgumentException(
                    "Could not parse date: " + e.Message + ". Input was: \"" + inputDate + "\"", e);
            }
            catch (ArgumentException e)
            {
                throw new ArgumentException(
                    "Could not parse date: " + e.Message + ". Input was: \"" + inputDate + "\"", e);
            }
        }
        #endregion

        #region AdjustTimezone
        /// <summary>
        /// Adjust the <paramref name="dateTime"/> object given according to the timezone specified in the <paramref name="dateInput"/>.
        /// </summary>
        /// <param name="dateTime">The date to alter</param>
        /// <param name="dateInput">The input date, in which the timezone can be found</param>
        /// <returns>An date altered according to the timezone</returns>
        private static DateTime AdjustTimezone(DateTime dateTime, string dateInput)
        {
            // We know that the timezones are always in the last part of the date input
            var parts = dateInput.Split(' ');
            var lastPart = parts[parts.Length - 1];

            // Convert timezones in older formats to [+-]dddd format.
            lastPart = Regex.Replace(lastPart, RegexOldTimezoneFormats, MatchEvaluator);

            // Find the timezone specification
            // Example: Fri, 21 Nov 1997 09:55:06 -0600
            // finds -0600
            var match = Regex.Match(lastPart, RegexNewTimezoneFormats);
            if (!match.Success) return dateTime;

            // We have found that the timezone is in +dddd or -dddd format
            // Add the number of hours and minutes to our found date
            var hours = int.Parse(match.Groups["hours"].Value);
            var minutes = int.Parse(match.Groups["minutes"].Value);

            var factor = match.Value[0] == '+' ? -1 : 1;

            dateTime = dateTime.AddHours(factor*hours);
            dateTime = dateTime.AddMinutes(factor*minutes);

            return dateTime;
            // A timezone of -0000 is the same as doing nothing
        }
        #endregion

        #region MatchEvaluator
        /// <summary>
        /// Convert timezones in older formats to [+-]dddd format.
        /// </summary>
        /// <param name="match">The match that was found</param>
        /// <returns>The string to replace the matched string with</returns>
        /// <remarks>
        ///
        /// RFC 2822: http://www.rfc-base.org/rfc-2822.html
        ///     
        ///     4.3. Obsolete Date and Time  
        /// 
        ///         The syntax for the obsolete date format allows a 2 digit year in the
        ///         date field and allows for a list of alphabetic time zone
        ///         specifications that were used in earlier versions of this standard.
        ///         It also permits comments and folding white space between many of the
        ///         tokens.
        ///     
        ///     	obs-day-of-week =       [CFWS] day-name [CFWS]
        ///     	obs-year        =       [CFWS] 2*DIGIT [CFWS]
        ///     	obs-month       =       CFWS month-name CFWS
        ///     	obs-day         =       [CFWS] 1*2DIGIT [CFWS]
        ///     	obs-hour        =       [CFWS] 2DIGIT [CFWS]
        ///     	obs-minute      =       [CFWS] 2DIGIT [CFWS]
        ///     	obs-second      =       [CFWS] 2DIGIT [CFWS]
        ///     	obs-zone        =       "UT" / "GMT" /          ; Universal Time
        ///     
        ///     	Resnick                     Standards Track                    [Page 31]
        ///     	RFC 2822                Internet Message Format               April 2001
        ///     
        ///     	                                                ; North American UT
        ///     	                                                ; offsets
        ///     	                        "EST" / "EDT" /         ; Eastern:  - 5/ - 4
        ///     	                        "CST" / "CDT" /         ; Central:  - 6/ - 5
        ///     	                        "MST" / "MDT" /         ; Mountain: - 7/ - 6
        ///     	                        "PST" / "PDT" /         ; Pacific:  - 8/ - 7
        ///     
        ///     	                        %d65-73 /               ; Military zones - "A"
        ///     	                        %d75-90 /               ; through "I" and "K"
        ///     	                        %d97-105 /              ; through "Z", both
        ///     	                        %d107-122               ; upper and lower case -- imported lower and upper
        ///
        /// </remarks>
        private static string MatchEvaluator(Match match)
        {
            if (!match.Success)
            {
                throw new ArgumentException("Match success are always true");
            }

            switch (match.Value)
            {
                // "A" through "I" and "a" through "i"
                // are equivalent to "+0100" through "+0900" respectively
                case "A":
                case "a":
                    return "+0100";
                case "B":
                case "b":
                    return "+0200";
                case "C":
                case "c":
                    return "+0300";
                case "D":
                case "d":
                    return "+0400";
                case "E":
                case "e":
                    return "+0500";
                case "F":
                case "f":
                    return "+0600";
                case "G":
                case "g":
                    return "+0700";
                case "H":
                case "h":
                    return "+0800";
                case "I":
                case "i":
                    return "+0900";

                // "K", "L", and "M" and "k", "l" and "m"
                // are equivalent to "+1000", "+1100", and "+1200" respectively
                case "K":
                case "k":
                    return "+1000";
                case "L":
                case "l":
                    return "+1100";
                case "M":
                case "m":
                    return "+1200";

                // "N" through "Y" and "n" through "y"
                // are equivalent to "-0100" through "-1200" respectively
                case "N":
                case "n":
                    return "-0100";
                case "O":
                case "o":
                    return "-0200";
                case "P":
                case "p":
                    return "-0300";
                case "Q":
                case "q":
                    return "-0400";
                case "R":
                case "r":
                    return "-0500";
                case "S":
                case "s":
                    return "-0600";
                case "T":
                case "t":
                    return "-0700";
                case "U":
                case "u":
                    return "-0800";
                case "V":
                case "v":
                    return "-0900";
                case "W":
                case "w":
                    return "-1000";
                case "X":
                case "x":
                    return "-1100";
                case "Y":
                case "y":
                    return "-1200";

                // "Z", "z", "UT" and "GMT"
                // is equivalent to "+0000"
                case "Z":
                case "z":
                case "UT":
                case "GMT":
                    return "+0000";

                // US time zones
                case "EDT":
                    return "-0400"; // EDT is semantically equivalent to -0400
                case "EST":
                    return "-0500"; // EST is semantically equivalent to -0500
                case "CDT":
                    return "-0500"; // CDT is semantically equivalent to -0500
                case "CST":
                    return "-0600"; // CST is semantically equivalent to -0600
                case "MDT":
                    return "-0600"; // MDT is semantically equivalent to -0600
                case "MST":
                    return "-0700"; // MST is semantically equivalent to -0700
                case "PDT":
                    return "-0700"; // PDT is semantically equivalent to -0700
                case "PST":
                    return "-0800"; // PST is semantically equivalent to -0800

                // EU time zones
                case "MSK":
                    return "+0400"; // MSK is semantically equivalent to +0400

                default:
                    throw new ArgumentException("Unexpected input");
            }
        }   
        #endregion

        #region ExtractDateTime
        /// <summary>
        /// Extracts the date and time parts from the <paramref name="dateInput"/>
        /// </summary>
        /// <param name="dateInput">The date input string, from which to extract the date and time parts</param>
        /// <returns>The extracted date part or <see langword="DateTime.MinValue"/> if <paramref name="dateInput"/> is not recognized as a valid date.</returns>
        /// <exception cref="ArgumentNullException">If <paramref name="dateInput"/> is <see langword="null"/></exception>
        private static DateTime ExtractDateTime(string dateInput)
        {
            if (dateInput == null)
                throw new ArgumentNullException("dateInput");

            // Matches the date and time part of a string
            // Given string example: Fri, 21 Nov 1997 09:55:06 -0600
            // Needs to find: 21 Nov 1997 09:55:06

            // Seconds does not need to be specified
            // Even though it is illigal, sometimes hours, minutes or seconds are only specified with one digit

            // Year with 2 or 4 digits (1922 or 22)
            const string year = @"(\d\d\d\d|\d\d)";

            // Time with one or two digits for hour and minute and optinal seconds (06:04:06 or 6:4:6 or 06:04 or 6:4)
            const string time = @"\d?\d:\d?\d(:\d?\d)?";

            // Correct format is 21 Nov 1997 09:55:06
            const string correctFormat = @"\d\d? .+ " + year + " " + time;

            // Some uses incorrect format: 2012-1-1 12:30
            const string incorrectFormat = year + @"-\d?\d-\d?\d " + time;

            // Some uses incorrect format: 08-May-2012 16:52:30 +0100
            const string correctFormatButWithDashes = @"\d\d?-[A-Za-z]{3}-" + year + " " + time;

            // We allow both correct and incorrect format
            const string joinedFormat = @"(" + correctFormat + ")|(" + incorrectFormat + ")|(" + correctFormatButWithDashes + ")";

            var match = Regex.Match(dateInput, joinedFormat);
            if (match.Success)
            {
                try
                {
                    return Convert.ToDateTime(match.Value, CultureInfo.InvariantCulture);
                }
                catch (FormatException)
                {
                }
            }

            //If there are some custom formats
            if (CustomDateTimeFormats == null) return DateTime.MinValue;
            //If there is a timezone at the end, remove it
            
            var strDate = dateInput.Trim();
            if (strDate.Contains(" ")) //Check contains a space before getting the last part to prevent accessing index -1
            {
                var parts = strDate.Split(' ');
                var lastPart = parts[parts.Length - 1];

                // Convert timezones in older formats to [+-]dddd format.
                lastPart = Regex.Replace(lastPart, RegexOldTimezoneFormats, MatchEvaluator);

                // Find the timezone specification
                // Example: Fri, 21 Nov 1997 09:55:06 -0600
                // finds -0600
                var timezoneMatch = Regex.Match(lastPart, RegexNewTimezoneFormats);
                if (timezoneMatch.Success)
                {
                    //This last part is a timezone, remove it
                    strDate = strDate.Substring(0, strDate.Length - parts[parts.Length - 1].Length).Trim(); //Use the length of the old last part
                }
            }

            //Try and parse it as one of the custom formats
            try
            {
                return DateTime.ParseExact(strDate, CustomDateTimeFormats, null, DateTimeStyles.None);
            }
            catch (FormatException)
            {
            }

            return DateTime.MinValue;
        }
        #endregion

        #region StripCommentsAndExcessWhitespace
        /// <summary>
        /// Strips and removes all comments and excessive whitespace from the string
        /// </summary>
        /// <param name="input">The input to strip from</param>
        /// <returns>The stripped string</returns>
        /// <exception cref="ArgumentNullException">If <paramref name="input"/> is <see langword="null"/></exception>
        private static string StripCommentsAndExcessWhitespace(string input)
        {
            if (input == null)
                throw new ArgumentNullException("input");

            // Strip out comments
            // Also strips out nested comments
            input = Regex.Replace(input, @"(\((?>\((?<C>)|\)(?<-C>)|.?)*(?(C)(?!))\))", "");

            // Reduce any whitespace character to one space only
            input = Regex.Replace(input, @"\s+", " ");

            // Remove all initial whitespace
            input = Regex.Replace(input, @"^\s+", "");

            // Remove all ending whitespace
            input = Regex.Replace(input, @"\s+$", "");

            // Remove spaces at colons
            // Example: 22: 33 : 44 => 22:33:44
            input = Regex.Replace(input, @" ?: ?", ":");

            return input;
        }
        #endregion

        #region FixSpecialCases
        /// <summary>
        /// Converts date time string in very wrong date time format:
        /// Tue Feb 18 10:23:30 2014 (MSK)
        /// to
        /// Feb 18 2014 10:23:30 MSK
        /// </summary>
        /// <param name="inputDate">The date to convert</param>
        /// <returns>The corrected string</returns>
        private static string FixSpecialCases(string inputDate)
        {
            const string weekDayPattern = "(?<weekDay>Mon|Tue|Wed|Thu|Fri|Sat|Sun)";
            const string monthPattern = @"(?<month>[A-Za-z]+)";
            const string dayPattern = @"(?<day>\d?\d)";
            const string yearPattern = @"(?<year>\d\d\d\d)";
            const string timePattern = @"(?<time>\d?\d:\d?\d(:\d?\d)?)";
            const string timeZonePattern = @"(?<timeZone>[A-Z]{3})";

            var incorrectFormat = String.Format(@"{0} +{1} +{2} +{3} +{4} +\({5}\)", weekDayPattern, monthPattern,
                dayPattern, timePattern, yearPattern, timeZonePattern);

            var match = Regex.Match(inputDate, incorrectFormat);
            if (!match.Success) return inputDate;
            var month = match.Groups["month"];
            var day = match.Groups["day"];
            var year = match.Groups["year"];
            var time = match.Groups["time"];
            var timeZone = match.Groups["timeZone"];
            return String.Format("{0} {1} {2} {3} {4}", day, month, year, time, timeZone);
        }
        #endregion
    }
}