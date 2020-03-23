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
	        catch (Exception)
	        {
				Logger.WriteToLog("Base64 deconding failed");
	            return new byte[0];
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
    }
}