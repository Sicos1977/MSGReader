using System.Text.RegularExpressions;

namespace MsgReader.Helpers
{
    /// <summary>
    /// This class contains helper functions for E-mail addresses
    /// </summary>
    internal static class EmailAddress
    {
        #region IsEmailAddressValid
        /// <summary>
        /// Return true when the E-mail address is valid
        /// </summary>
        /// <param name="emailAddress"></param>
        /// <returns></returns>
        public static bool IsEmailAddressValid(string emailAddress)
        {
            if (string.IsNullOrEmpty(emailAddress))
                return false;

            var regex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.IgnoreCase);
            var matches = regex.Matches(emailAddress);

            return matches.Count == 1;
        }
        #endregion

        #region RemoveSingleQuotes
        /// <summary>
        /// Removes trailing en ending single quotes from an E-mail address when they exist
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        public static string RemoveSingleQuotes(string email)
        {
            if (string.IsNullOrEmpty(email))
                return string.Empty;

            if (email.StartsWith("'"))
                email = email.Substring(1, email.Length - 1);

            if (email.EndsWith("'"))
                email = email.Substring(0, email.Length - 1);

            return email;
        }
        #endregion
    }
}
