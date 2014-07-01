using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DocumentServices.Modules.Readers.MsgReader.Mime.Decode
{
	/// <summary>
	/// Class used to decode RFC 2822 Date header fields.
	/// </summary>
	internal static class Rfc2822DateTime
    {
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
	            throw new ArgumentException("Could not parse date: " + e.Message + ". Input was: \"" + inputDate + "\"", e);
	        }
	        catch (ArgumentException e)
	        {
	            throw new ArgumentException("Could not parse date: " + e.Message + ". Input was: \"" + inputDate + "\"", e);
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
	        lastPart = Regex.Replace(lastPart, @"UT|GMT|EST|EDT|CST|CDT|MST|MDT|PST|MSK|PDT|[A-I]|[K-Y]|Z", MatchEvaluator);

	        // Find the timezone specification
	        // Example: Fri, 21 Nov 1997 09:55:06 -0600
	        // finds -0600
	        var match = Regex.Match(lastPart, @"[\+-](?<hours>\d\d)(?<minutes>\d\d)");
	        if (match.Success)
	        {
	            // We have found that the timezone is in +dddd or -dddd format
	            // Add the number of hours and minutes to our found date
	            var hours = int.Parse(match.Groups["hours"].Value);
	            var minutes = int.Parse(match.Groups["minutes"].Value);

	            var factor = match.Value[0] == '+' ? -1 : 1;

	            dateTime = dateTime.AddHours(factor*hours);
	            dateTime = dateTime.AddMinutes(factor*minutes);

	            return dateTime;
	        }

	        // A timezone of -0000 is the same as doing nothing
	        return dateTime;
	    }
	    #endregion

        #region MatchEvaluator
        /// <summary>
	    /// Convert timezones in older formats to [+-]dddd format.
	    /// </summary>
	    /// <param name="match">The match that was found</param>
	    /// <returns>The string to replace the matched string with</returns>
	    private static string MatchEvaluator(Match match)
	    {
	        if (!match.Success)
	        {
	            throw new ArgumentException("Match success are always true");
	        }

	        switch (match.Value)
	        {
	                // "A" through "I"
	                // are equivalent to "+0100" through "+0900" respectively
	            case "A":
	                return "+0100";
	            case "B":
	                return "+0200";
	            case "C":
	                return "+0300";
	            case "D":
	                return "+0400";
	            case "E":
	                return "+0500";
	            case "F":
	                return "+0600";
	            case "G":
	                return "+0700";
	            case "H":
	                return "+0800";
	            case "I":
	                return "+0900";

	                // "K", "L", and "M"
	                // are equivalent to "+1000", "+1100", and "+1200" respectively
	            case "K":
	                return "+1000";
	            case "L":
	                return "+1100";
	            case "M":
	                return "+1200";

	                // "N" through "Y"
	                // are equivalent to "-0100" through "-1200" respectively
	            case "N":
	                return "-0100";
	            case "O":
	                return "-0200";
	            case "P":
	                return "-0300";
	            case "Q":
	                return "-0400";
	            case "R":
	                return "-0500";
	            case "S":
	                return "-0600";
	            case "T":
	                return "-0700";
	            case "U":
	                return "-0800";
	            case "V":
	                return "-0900";
	            case "W":
	                return "-1000";
	            case "X":
	                return "-1100";
	            case "Y":
	                return "-1200";

	                // "Z", "UT" and "GMT"
	                // is equivalent to "+0000"
	            case "Z":
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
	        const string joinedFormat =
	            @"(" + correctFormat + ")|(" + incorrectFormat + ")|(" + correctFormatButWithDashes + ")";

	        var match = Regex.Match(dateInput, joinedFormat);
	        if (match.Success)
	            return Convert.ToDateTime(match.Value, CultureInfo.InvariantCulture);

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