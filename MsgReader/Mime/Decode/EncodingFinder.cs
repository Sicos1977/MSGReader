using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace MsgReader.Mime.Decode
{
	/// <summary>
	/// Utility class used by OpenPop for mapping from a characterSet to an <see cref="Encoding"/>.<br/>
	/// <br/>
	/// The functionality of the class can be altered by adding mappings
	/// using <see cref="AddMapping"/> and by adding a <see cref="FallbackDecoder"/>.<br/>
	/// <br/>
	/// Given a characterSet, it will try to find the Encoding as follows:
	/// <list type="number">
	///     <item>
	///         <description>If a mapping for the characterSet was added, use the specified Encoding from there. Mappings can be added using <see cref="AddMapping"/>.</description>
	///     </item>
	///     <item>
	///         <description>Try to parse the characterSet and look it up using <see cref="Encoding.GetEncoding(int)"/> for codepages or <see cref="Encoding.GetEncoding(string)"/> for named encodings.</description>
	///     </item>
	///     <item>
	///         <description>If an encoding is not found yet, use the <see cref="FallbackDecoder"/> if defined. The <see cref="FallbackDecoder"/> is user defined.</description>
	///     </item>
	/// </list>
	/// </summary>
	public static class EncodingFinder
    {
        #region Proprties
        /// <summary>
		/// Delegate that is used when the EncodingFinder is unable to find an encoding by
		/// using the <see cref="EncodingFinder.EncodingMap"/> or general code.<br/>
		/// This is used as a last resort and can be used for setting a default encoding or
		/// for finding an encoding on runtime for some <paramref name="characterSet"/>.
		/// </summary>
		/// <param name="characterSet">The character set to find an encoding for.</param>
		/// <returns>An encoding for the <paramref name="characterSet"/> or <see langword="null"/> if none could be found.</returns>
		public delegate Encoding FallbackDecoderDelegate(string characterSet);

		/// <summary>
		/// Last resort decoder.
		/// </summary>
		public static FallbackDecoderDelegate FallbackDecoder { private get; set; }

		/// <summary>
		/// Mapping from charactersets to encodings.
		/// </summary>
		private static Dictionary<string, Encoding> EncodingMap { get; set; }
        #endregion

	    #region Constructor
	    /// <summary>
	    /// Initialize the EncodingFinder
	    /// </summary>
	    static EncodingFinder()
	    {
	        Reset();
	    }
	    #endregion

        #region Reset
        /// <summary>
	    /// Used to reset this static class to facilite isolated unit testing.
	    /// </summary>
	    internal static void Reset()
	    {
	        EncodingMap = new Dictionary<string, Encoding>();
	        FallbackDecoder = null;

	        // Some emails incorrectly specify the encoding as utf8, but it should have been utf-8.
	        AddMapping("utf8", Encoding.UTF8);
	        AddMapping("binary", Encoding.ASCII);
	    }
	    #endregion

        #region FindEncoding
        /// <summary>
	    /// Parses a character set into an encoding.
	    /// </summary>
	    /// <param name="characterSet">The character set to parse</param>
	    /// <returns>An encoding which corresponds to the character set</returns>
	    /// <exception cref="ArgumentNullException">If <paramref name="characterSet"/> is <see langword="null"/></exception>
	    internal static Encoding FindEncoding(string characterSet)
	    {
	        if (characterSet == null)
	            throw new ArgumentNullException("characterSet");

	        var charSetUpper = characterSet.ToUpperInvariant();

	        // Check if the characterSet is explicitly mapped to an encoding
	        if (EncodingMap.ContainsKey(charSetUpper))
	            return EncodingMap[charSetUpper];

	        // Try to generally find the encoding
	        try
	        {
	            if (!charSetUpper.Contains("WINDOWS") && !charSetUpper.Contains("CP"))
	                return Encoding.GetEncoding(characterSet);
	            // It seems the characterSet contains an codepage value, which we should use to parse the encoding
	            charSetUpper = charSetUpper.Replace("CP", ""); // Remove cp
	            charSetUpper = charSetUpper.Replace("WINDOWS", ""); // Remove windows
	            charSetUpper = charSetUpper.Replace("-", ""); // Remove - which could be used as cp-1554

	            // Now we hope the only thing left in the characterSet is numbers.
	            var codepageNumber = int.Parse(charSetUpper, CultureInfo.InvariantCulture);

	            return Encoding.GetEncoding(codepageNumber);

	            // It seems there is no codepage value in the characterSet. It must be a named encoding
	        }
	        catch (ArgumentException)
	        {
	            // The encoding could not be found generally. 
	            // Try to use the FallbackDecoder if it is defined.

	            // Check if it is defined
	            if (FallbackDecoder == null)
	                throw; // It was not defined - throw catched exception

	            // Use the FallbackDecoder
	            var fallbackDecoderResult = FallbackDecoder(characterSet);

	            // Check if the FallbackDecoder had a solution
	            if (fallbackDecoderResult != null)
	                return fallbackDecoderResult;

	            // If no solution was found, throw catched exception
	            throw;
	        }
	    }
	    #endregion

        #region AddMapping
        /// <summary>
	    /// Puts a mapping from <paramref name="characterSet"/> to <paramref name="encoding"/>
	    /// into the <see cref="EncodingFinder"/>'s internal mapping Dictionary.
	    /// </summary>
	    /// <param name="characterSet">The string that maps to the <paramref name="encoding"/></param>
	    /// <param name="encoding">The <see cref="Encoding"/> that should be mapped from <paramref name="characterSet"/></param>
	    /// <exception cref="ArgumentNullException">If <paramref name="characterSet"/> is <see langword="null"/></exception>
	    /// <exception cref="ArgumentNullException">If <paramref name="encoding"/> is <see langword="null"/></exception>
	    public static void AddMapping(string characterSet, Encoding encoding)
	    {
	        if (characterSet == null)
	            throw new ArgumentNullException("characterSet");

	        if (encoding == null)
	            throw new ArgumentNullException("encoding");

	        // Add the mapping using uppercase
	        EncodingMap.Add(characterSet.ToUpperInvariant(), encoding);
	    }
	    #endregion
	}
}
