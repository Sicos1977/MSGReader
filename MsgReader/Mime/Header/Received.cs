using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using MsgReader.Mime.Decode;

namespace MsgReader.Mime.Header
{
	/// <summary>
	/// Class that hold information about one "Received:" header line.<br/>
	/// <br/>
	/// Visit these RFCs for more information:<br/>
	/// <see href="http://tools.ietf.org/html/rfc5321#section-4.4">RFC 5321 section 4.4</see><br/>
	/// <see href="http://tools.ietf.org/html/rfc4021#section-3.6.7">RFC 4021 section 3.6.7</see><br/>
	/// <see href="http://tools.ietf.org/html/rfc2822#section-3.6.7">RFC 2822 section 3.6.7</see><br/>
	/// <see href="http://tools.ietf.org/html/rfc2821#section-4.4">RFC 2821 section 4.4</see><br/>
	/// </summary>
	public class Received
	{
	    #region Properties
	    /// <summary>
	    /// The date of this received line.
	    /// Is <see cref="DateTime.MinValue"/> if not present in the received header line.
	    /// </summary>
	    public DateTime Date { get; private set; }

	    /// <summary>
	    /// A dictionary that contains the names and values of the
	    /// received header line.<br/>
	    /// If the received header is invalid and contained one name
	    /// multiple times, the first one is used and the rest is ignored.
	    /// </summary>
	    /// <example>
	    /// If the header lines looks like:
	    /// <code>
	    /// from sending.com (localMachine [127.0.0.1]) by test.net (Postfix)
	    /// </code>
	    /// then the dictionary will contain two keys: "from" and "by" with the values
	    /// "sending.com (localMachine [127.0.0.1])" and "test.net (Postfix)".
	    /// </example>
	    public Dictionary<string, string> Names { get; private set; }

	    /// <summary>
	    /// The raw input string that was parsed into this class.
	    /// </summary>
	    public string Raw { get; private set; }
	    #endregion

        #region Received
        /// <summary>
	    /// Parses a Received header value.
	    /// </summary>
	    /// <param name="headerValue">The value for the header to be parsed</param>
	    /// <exception cref="ArgumentNullException"><exception cref="ArgumentNullException">If <paramref name="headerValue"/> is <see langword="null"/></exception></exception>
	    public Received(string headerValue)
	    {
	        if (headerValue == null)
	            throw new ArgumentNullException("headerValue");

	        // Remember the raw input if someone whishes to use it
	        Raw = headerValue;

	        // Default Date value
	        Date = DateTime.MinValue;

	        // The date part is the last part of the string, and is preceeded by a semicolon
	        // Some emails forgets to specify the date, therefore we need to check if it is there
	        if (headerValue.Contains(";"))
	        {
	            var datePart = headerValue.Substring(headerValue.LastIndexOf(";", StringComparison.Ordinal) + 1);
	            Date = Rfc2822DateTime.StringToDate(datePart);
	        }

	        Names = ParseDictionary(headerValue);
	    }
	    #endregion

        #region ParseDictionary
        /// <summary>
	    /// Parses the Received header name-value-list into a dictionary.
	    /// </summary>
	    /// <param name="headerValue">The full header value for the Received header</param>
	    /// <returns>A dictionary where the name-value-list has been parsed into</returns>
	    private static Dictionary<string, string> ParseDictionary(string headerValue)
	    {
	        var dictionary = new Dictionary<string, string>();

	        // Remove the date part from the full headerValue if it is present
	        var headerValueWithoutDate = headerValue;
	        if (headerValue.Contains(";"))
	            headerValueWithoutDate = headerValue.Substring(0, headerValue.LastIndexOf(";", StringComparison.Ordinal));

	        // Reduce any whitespace character to one space only
	        headerValueWithoutDate = Regex.Replace(headerValueWithoutDate, @"\s+", " ");

	        // The regex below should capture the following:
	        // The name consists of non-whitespace characters followed by a whitespace and then the value follows.
	        // There are multiple cases for the value part:
	        //   1: Value is just some characters not including any whitespace
	        //   2: Value is some characters, a whitespace followed by an unlimited number of
	        //      parenthesized values which can contain whitespaces, each delimited by whitespace
	        //
	        // Cheat sheet for regex:
	        // \s means every whitespace character
	        // [^\s] means every character except whitespace characters
	        // +? is a non-greedy equivalent of +
	        const string pattern = @"(?<name>[^\s]+)\s(?<value>[^\s]+(\s\(.+?\))*)";

	        // Find each match in the string
	        var matches = Regex.Matches(headerValueWithoutDate, pattern);
	        foreach (Match match in matches)
	        {
	            // Add the name and value part found in the matched result to the dictionary
	            var name = match.Groups["name"].Value;
	            var value = match.Groups["value"].Value;

	            // Check if the name is really a comment.
	            // In this case, the first entry in the header value
	            // is a comment
	            if (name.StartsWith("("))
	                continue;

	            // Only add the first name pair
	            // All subsequent pairs are ignored, as they are invalid anyway
	            if (!dictionary.ContainsKey(name))
	                dictionary.Add(name, value);
	        }

	        return dictionary;
	    }
	    #endregion
	}
}