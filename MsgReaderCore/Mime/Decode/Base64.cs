using System;
using System.Text;
using MsgReader.Helpers;

namespace MsgReader.Mime.Decode
{
	/// <summary>
	/// Utility class for dealing with Base64 encoded strings
	/// </summary>
	internal static class Base64
    {
        #region Decode
	    /// <summary>
	    /// Decodes a base64 encoded string into the bytes it describes
	    /// </summary>
	    /// <param name="base64Encoded">The string to decode</param>
	    /// <returns>A byte array that the base64 string described</returns>
	    public static byte[] Decode(string base64Encoded)
        {
	        try
	        {
                return Convert.FromBase64String(base64Encoded);
            }
	        catch (Exception exception)
	        {
				Logger.WriteToLog($"Base64 decoding failed, error: {ExceptionHelpers.GetInnerException(exception)}, trying to fix it by removing invalid chars");
                var base64Chars = RemoveInvalidBase64Chars(base64Encoded);

                try
                {
                    return Convert.FromBase64String(base64Chars);
                }
                catch (Exception)
                {
                    Logger.WriteToLog("Base64 decoding still failed returning empty byte array");
                    return new byte[0];
                }
	        }
	    }

	    /// <summary>
	    /// Decodes a Base64 encoded string using a specified <see cref="System.Text.Encoding"/> 
	    /// </summary>
	    /// <param name="base64Encoded">Source string to decode</param>
	    /// <param name="encoding">The encoding to use for the decoded byte array that <paramref name="base64Encoded"/> describes</param>
	    /// <returns>A decoded string</returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="base64Encoded"/> or <paramref name="encoding"/> is <see langword="null"/></exception>
	    /// <exception cref="FormatException">If <paramref name="base64Encoded"/> is not a valid base64 encoded string</exception>
	    public static string Decode(string base64Encoded, Encoding encoding)
	    {
	        if (base64Encoded == null)
	            throw new ArgumentNullException(nameof(base64Encoded));

	        if (encoding == null)
	            throw new ArgumentNullException(nameof(encoding));

	        try
	        {
                return encoding.GetString(Decode(base64Encoded));
            }
            catch (FormatException)
	        {
	            return string.Empty;
	        }
	    }
	    #endregion

        #region RemoveInvalidBase64Chars
        /// <summary>
        /// Removes everything from a base64 string that does not belong in it
        /// </summary>
        /// <param name="base64Encoded"></param>
        /// <returns></returns>
        private static string RemoveInvalidBase64Chars(string base64Encoded)
        {
            var result = new StringBuilder();

            foreach (var chr in base64Encoded)
            {
                var val = (int) chr;
                if (val >= 65 && val <= 90 ||	// 'A'..'Z'
                    val >= 97 && val <= 122 ||	// 'a'..'z'
                    val >= 48 && val <= 57 ||   // '0'..'9'
                    val == 43 || val == 47)     // '+' and '/'
                    result.Append(chr);
                else
                    Logger.WriteToLog($"Removing char with value '{val}' from base64 string");
            }

            // Add the correct amount of padding to cleaned string
            var str = result.ToString();

            // Remove padding first if there is any
            str = str.TrimEnd('=');

            var padding = str.Length % 4;
            // Add correct amount of padding
            str += padding % 2 > 0 ? "=" : "==";

            return str;
        }
        #endregion
    }
}