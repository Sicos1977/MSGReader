// name       : StringsBase.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2009.02.19
// language   : c#
// environment: .NET 3.5
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Resources;

namespace Itenso.Sys
{
    /// <summary>
    ///     Provides some helper functionality to keep resource handling in a
    ///     namespace as simple and uniform as possible.
    /// </summary>
    /// <remarks>
    ///     intended to be used by a singleton class per namespace, which should
    ///     commonly be named <c>Strings</c>. This singleton instance should
    ///     provide access to resources via strongly typed properties, thus
    ///     avoiding coding errors with misspelled resource identifiers.
    /// </remarks>
    public abstract class StringsBase
    {
        /// <summary>
        ///     Formats the given format-string with the invariant culture and the
        ///     given arguments.
        /// </summary>
        /// <param name="format">the string to format</param>
        /// <param name="args">the arguments to fill in</param>
        /// <returns>the formatted string</returns>
        protected static string Format(string format, params object[] args)
        {
            return StringTool.FormatSafeInvariant(format, args);
        } // Format

        /// <summary>
        ///     Creates a <c>ResourceManager</c> instance for the given type, loading its resources
        ///     from the type's full name, suffixed with '<c>.resx</c>'.
        /// </summary>
        /// <param name="singletonType">the type of the singleton</param>
        /// <returns>a <c>ResourceManager</c> for loading the given type's resources</returns>
        protected static ResourceManager NewInst(Type singletonType)
        {
            if (singletonType == null)
                throw new ArgumentNullException(nameof(singletonType));
            if (string.IsNullOrEmpty(singletonType.FullName))
                throw new ArgumentException("singletonType");
            return new ResourceManager(singletonType.FullName, singletonType.Assembly);
        } // NewInst
    }
}