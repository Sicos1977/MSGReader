// -- FILE ------------------------------------------------------------------
// name       : ArgumentCheck.cs
// project    : System Framelet
// created    : Leon Poyyayil - 2005.05.02
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;

namespace Itenso.Sys
{
    // ----------------------------------------------------------------------
    /// <summary>
    ///     Some helper methods for common checks on arguments.
    /// </summary>
    public static class ArgumentCheck
    {
        // ----------------------------------------------------------------------
        /// <summary>
        ///     Checks the given value and returns it.
        /// </summary>
        /// <param name="value">the value to check</param>
        /// <param name="name">the name for the <c>ArgumentException</c></param>
        /// <returns>the trimmed value</returns>
        /// <exception cref="ArgumentNullException">in case the given value is null</exception>
        /// <exception cref="ArgumentException">in case the trimmed given value is empty</exception>
        public static string NonemptyTrimmedString(string value, string name)
        {
            return NonemptyTrimmedString(value, Strings.ArgumentMayNotBeEmpty, name);
        } // NonemptyTrimmedString

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Checks the given value and returns it.
        /// </summary>
        /// <param name="value">the value to check</param>
        /// <param name="exceptionMessage">
        ///     the message in the <see cref="ArgumentException" />
        ///     in case the given value is empty
        /// </param>
        /// <param name="name">the name for the <c>ArgumentException</c></param>
        /// <returns>the trimmed value</returns>
        /// <exception cref="ArgumentNullException">in case the given value is null</exception>
        /// <exception cref="ArgumentException">in case the trimmed given value is empty</exception>
        public static string NonemptyTrimmedString(string value, string exceptionMessage, string name)
        {
            if (value == null)
                throw new ArgumentNullException(name);
            var trimmed = value.Trim();
            if (trimmed.Length == 0)
                throw new ArgumentException(exceptionMessage, name);
            return trimmed;
        } // NonemptyTrimmedString
    } // class ArgumentCheck
} // namespace Itenso.Sys
// -- EOF -------------------------------------------------------------------