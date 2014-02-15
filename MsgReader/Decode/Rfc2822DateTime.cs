using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace DocumentServices.Modules.Readers.MsgReader.Decode
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
		/// <exception cref="ArgumentNullException"><exception cref="ArgumentNullException">If <paramref name="inputDate"/> is <see langword="null"/></exception></exception>
		/// <exception cref="ArgumentException">If the <paramref name="inputDate"/> could not be parsed into a <see cref="DateTime"/> object</exception>
		public static DateTime StringToDate(string inputDate)
		{
			if(inputDate == null)
				throw new ArgumentNullException("inputDate");

			// Old date specification allows comments and a lot of whitespace
			inputDate = StripCommentsAndExcessWhitespace(inputDate);

			try
			{
                // Extract the DateTime
                var dateTime = ExtractDateTime(inputDate);

				// If a day-name is specified in the inputDate string, check if it fits with the date
				ValidateDayNameIfAny(dateTime, inputDate);

				// Convert the date into UTC
				dateTime = new DateTime(dateTime.Ticks, DateTimeKind.Utc);

				// Adjust according to the time zone
				dateTime = AdjustTimezone(dateTime, inputDate);
				
				// Return the parsed date
				return dateTime;
			}
			catch (FormatException e)	// Convert.ToDateTime() Failure
			{
				throw new ArgumentException("Could not parse date: " + e.Message + ". Input was: \"" + inputDate + "\"", e);
			}
			catch (ArgumentException)
			{
			    return DateTime.Now;
				//throw new ArgumentException("Could not parse date: " + e.Message + ". Input was: \"" + inputDate + "\"", e);
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
		/// <exception cref="ArgumentException">If no timezone was found in <paramref name="dateInput"/></exception>
		private static DateTime AdjustTimezone(DateTime dateTime, string dateInput)
		{
			// We know that the timezones are always in the last part of the date input
			var parts = dateInput.Split(' ');
			var lastPart = parts[parts.Length - 1];

			// Convert timezones in older formats to [+-]dddd format.
			lastPart = Regex.Replace(lastPart, @"UT|GMT|EST|EDT|CST|CDT|MST|MDT|PST|PDT|[A-I]|[K-Y]|Z", MatchEvaluator);

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
				case "A": return "+0100";
				case "B": return "+0200";
				case "C": return "+0300";
				case "D": return "+0400";
				case "E": return "+0500";
				case "F": return "+0600";
				case "G": return "+0700";
				case "H": return "+0800";
				case "I": return "+0900";

				// "K", "L", and "M"
				// are equivalent to "+1000", "+1100", and "+1200" respectively
				case "K": return "+1000";
				case "L": return "+1100";
				case "M": return "+1200";

				// "N" through "Y"
				// are equivalent to "-0100" through "-1200" respectively
				case "N": return "-0100";
				case "O": return "-0200";
				case "P": return "-0300";
				case "Q": return "-0400";
				case "R": return "-0500";
				case "S": return "-0600";
				case "T": return "-0700";
				case "U": return "-0800";
				case "V": return "-0900";
				case "W": return "-1000";
				case "X": return "-1100";
				case "Y": return "-1200";

				// "Z", "UT" and "GMT"
				// is equivalent to "+0000"
				case "Z":
				case "UT":
				case "GMT":
					return "+0000";

				// US time zones
				case "EDT": return "-0400"; // EDT is semantically equivalent to -0400
				case "EST": return "-0500"; // EST is semantically equivalent to -0500
				case "CDT": return "-0500"; // CDT is semantically equivalent to -0500
				case "CST": return "-0600"; // CST is semantically equivalent to -0600
				case "MDT": return "-0600"; // MDT is semantically equivalent to -0600
				case "MST": return "-0700"; // MST is semantically equivalent to -0700
				case "PDT": return "-0700"; // PDT is semantically equivalent to -0700
				case "PST": return "-0800"; // PST is semantically equivalent to -0800

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
		private static DateTime ExtractDateTime(string dateInput)
		{
			// Matches the date and time part of a string
			// Example: Fri, 21 Nov 1997 09:55:06 -0600
			// Finds: 21 Nov 1997 09:55:06
			// Seconds does not need to be specified
			// Even though it is illigal, sometimes hours, minutes or seconds are only specified with one digit
			Match match = Regex.Match(dateInput, @"\d\d? .+ (\d\d\d\d|\d\d) \d?\d:\d?\d(:\d?\d)?");
			if(match.Success)
			{
                return Convert.ToDateTime(match.Value, CultureInfo.InvariantCulture);
			}

			return DateTime.MinValue;
		}
        #endregion

        #region ValidateDayNameIfAny
        /// <summary>
		/// Validates that the given <paramref name="dateTime"/> agrees with a day-name specified
		/// in <paramref name="dateInput"/>.
		/// </summary>
		/// <param name="dateTime">The time to check</param>
		/// <param name="dateInput">The date input to extract the day-name from</param>
		/// <exception cref="ArgumentException">If <paramref name="dateTime"/> and <paramref name="dateInput"/> does not agree on the day</exception>
		private static void ValidateDayNameIfAny(DateTime dateTime, string dateInput)
		{
			// Check if there is a day name in front of the date
			// Example: Fri, 21 Nov 1997 09:55:06 -0600
			if (dateInput.Length >= 4 && dateInput[3] == ',')
			{
				string dayName = dateInput.Substring(0, 3);

				// If a dayName was specified. Check that the dateTime and the dayName
				// agrees on which day it is
				// This is just a failure-check and could be left out
				if ((dateTime.DayOfWeek == DayOfWeek.Monday    && !dayName.Equals("Mon")) ||
					(dateTime.DayOfWeek == DayOfWeek.Tuesday   && !dayName.Equals("Tue")) ||
					(dateTime.DayOfWeek == DayOfWeek.Wednesday && !dayName.Equals("Wed")) ||
					(dateTime.DayOfWeek == DayOfWeek.Thursday  && !dayName.Equals("Thu")) ||
					(dateTime.DayOfWeek == DayOfWeek.Friday    && !dayName.Equals("Fri")) ||
					(dateTime.DayOfWeek == DayOfWeek.Saturday  && !dayName.Equals("Sat")) ||
					(dateTime.DayOfWeek == DayOfWeek.Sunday    && !dayName.Equals("Sun")))
				{
                    throw new ArgumentException("Day-name does not correspond to the weekday of the date: " + dateInput);					
				}
			}

			// If no day name was found no checks can be made
		}
        #endregion

        #region StripCommentsAndExcessWhitespace
        /// <summary>
		/// Strips and removes all comments and excessive whitespace from the string
		/// </summary>
		/// <param name="input">The input to strip from</param>
		/// <returns>The stripped string</returns>
		private static string StripCommentsAndExcessWhitespace(string input)
		{
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
    }
}