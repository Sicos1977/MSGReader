// name       : CollectionTool.cs
// project    : System Framelet
// created    : Leon Poyyayil - 2005.05.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;
using System.Reflection;
using System.Text;

namespace Itenso.Sys.Collection
{
    /// <summary>
    ///     Some utility methods for collections.
    /// </summary>
    /// <remarks>
    ///     Just a container for some static methods which make life somewhat easier.
    /// </remarks>
    public static class CollectionTool
    {
        public static bool HaveSameContents(IEnumerable left, IEnumerable right)
        {
            var equal = left == right;
            if (!equal)
                if (left != null && right != null)
                {
                    var otherItems = right.GetEnumerator();
                    equal = true;
                    foreach (var item in left)
                        if (otherItems.MoveNext())
                        {
                            var otherItem = otherItems.Current;
                            if (item != otherItem && (item == null || !item.Equals(otherItem)))
                            {
                                equal = false;
                                break;
                            }
                        }
                        else
                        {
                            // the other enumeration has less objects
                            equal = false;
                            break;
                        }
                    if (equal && otherItems.MoveNext())
                        equal = false;
                }
            return equal;
        } // HaveSameContents

        public static bool AreEqual(IEnumerable enumerable, object obj)
        {
            var equal = enumerable == obj;
// ReSharper disable PossibleMultipleEnumeration
            if (!equal && enumerable != null && obj != null && enumerable.GetType() == obj.GetType())
                equal = HaveSameContents(enumerable, obj as IEnumerable);
            return equal;
        } // AreEqual

        public static int AddHashCode(int hash, object obj)
        {
            var combinedHash = obj != null ? obj.GetHashCode() : 0;
            if (hash != 0) // perform this check to prevent FxCop warning 'op could overflow'
                combinedHash += hash * 31;
            return combinedHash;
        } // AddHashCode

        public static int AddHashCode(int hash, int objHash)
        {
            var combinedHash = objHash;
            if (hash != 0) // perform this check to prevent FxCop warning 'op could overflow'
                combinedHash += hash * 31;
            return combinedHash;
        } // AddHashCode

        public static int ComputeHashCode(IEnumerable enumerable)
        {
            var hash = 1;
            if (enumerable == null)
                throw new ArgumentNullException("enumerable");
            foreach (var item in enumerable)
                hash = hash * 31 + (item != null ? item.GetHashCode() : 0);
            return hash;
        } // ComputeHashCode

        public static string ToString(IEnumerable enumerable)
        {
            return ToString(enumerable, "[", "]", ",", "null");
        } // ToString

        public static string ToString(IEnumerable enumerable, string delimiterText)
        {
            return ToString(enumerable, string.Empty, string.Empty, delimiterText, string.Empty);
        } // ToString

        /// <summary>
        ///     conventiently concatenates the given items to a string for debugging purposes.
        /// </summary>
        /// <remarks>
        ///     the whole collection is embraced with square brackets and the individual items
        ///     are separated by a comma. null items will be displayed as 'null' instead of the
        ///     empty string.
        /// </remarks>
        /// <param name="enumerable">the collection of items to print</param>
        /// <param name="startText">the starting text</param>
        /// <param name="endText">the ending textrint</param>
        /// <param name="delimiterText">the item delimiter text</param>
        /// <param name="undefinedValueText">text for undefined values</param>
        /// <returns>a concatenation of the string representations of all the items</returns>
        public static string ToString(IEnumerable enumerable, string startText, string endText, string delimiterText,
            string undefinedValueText)
        {
            if (enumerable == null)
                throw new ArgumentNullException("enumerable");
            var str = new StringBuilder(startText);
            var first = true;
            foreach (var obj in enumerable)
            {
                if (obj == null && string.IsNullOrEmpty(undefinedValueText))
                    continue;

                if (first)
                    first = false;
                else
                    str.Append(delimiterText);
                if (obj == null)
                {
                    str.Append(undefinedValueText);
                }
                else if (obj is DictionaryEntry)
                {
                    var mapEntry = (DictionaryEntry) obj;
                    str.Append(mapEntry.Key);
                    str.Append("=");
                    str.Append(mapEntry.Value == null ? undefinedValueText : mapEntry.Value.ToString());
                }
                else
                {
                    str.Append(obj);
                }
            }
            str.Append(endText);

            return str.ToString();
        } // ToString

        public static string EnumValuesToString(Type enumType)
        {
            return EnumValuesToString(enumType, "[", "]", "|");
        } // EnumValuesToString

        public static string EnumValuesToString(Type enumType, string delimiterText)
        {
            return EnumValuesToString(enumType, string.Empty, string.Empty, delimiterText);
        } // EnumValuesToString

        public static string EnumValuesToString(Type enumType, string startText, string endText, string delimiterText)
        {
            if (enumType == null)
                throw new ArgumentNullException("enumType");
            var str = new StringBuilder(startText);
            var fields = enumType.GetFields(BindingFlags.Public | BindingFlags.Static);
            for (var i = 0; i < fields.Length; i++)
            {
                if (i > 0)
                    str.Append(delimiterText);
                str.Append(fields[i].Name);
            }
            str.Append(endText);

            return str.ToString();
        } // EnumValuesToString

        public static int ParseEnumValue(Type enumType, string value, bool ignoreCase)
        {
            if (enumType == null)
                throw new ArgumentNullException("enumType");
            try
            {
                return (int) Enum.Parse(enumType, value, ignoreCase);
            }
            catch (ArgumentException)
            {
                try
                {
                    throw new ArgumentException(
                        Strings.CollectionToolInvalidEnum(value, enumType.Name, EnumValuesToString(enumType)));
                }
                catch (FormatException)
                {
                    // EXC: ignore, should not happen with a coded format string
                    return 0;
                }
            }
        } // ParseEnumValue
    }
}