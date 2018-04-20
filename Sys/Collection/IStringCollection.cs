// -- FILE ------------------------------------------------------------------
// name       : IStringCollection.cs
// project    : System Framelet
// created    : Leon Poyyayil - 2005.05.02
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Collections;

namespace Itenso.Sys.Collection
{
    // ------------------------------------------------------------------------
    /// <summary>
    ///     A simple immutable storage utility to hold multiple strings.
    /// </summary>
    public interface IStringCollection : IEnumerable
    {
        // ----------------------------------------------------------------------
        /// <summary>
        ///     Access to the number of items.
        /// </summary>
        /// <value>the number of items in the collection</value>
        int Count { get; }

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Index access to the items of this collection.
        /// </summary>
        /// <param name="index">the index of the item to retrieve</param>
        /// <returns>the item at the given position</returns>
        string this[int index] { get; }

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Copies this collections items to the given array.
        /// </summary>
        /// <param name="array">the target array</param>
        /// <param name="index">the target index</param>
        void CopyTo(string[] array, int index);

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Locates the given string in this collection.
        /// </summary>
        /// <param name="test">the string to search</param>
        /// <returns>the position of the given string in this collection or -1 if not found</returns>
        int IndexOf(string test);

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Tests whether the given string is present in this collection.
        /// </summary>
        /// <param name="test">the string to search</param>
        /// <returns>true if this collection contains such a string, false otherwise</returns>
        bool Contains(string test);

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Formats the lists items as a comma separated string without any special
        ///     quoting (e.g. if the items contain commas themselves ...)
        /// </summary>
        /// <returns>a string with all items separated by commas</returns>
        string FormatCommaSeparated();
    } // interface IStringCollection
} // namespace Itenso.Sys.Collection
// -- EOF -------------------------------------------------------------------