// name       : StringCollection.cs
// project    : System Framelet
// created    : Leon Poyyayil - 2005.04.27
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;
using System.Text;

namespace Itenso.Sys.Collection
{
    /// <summary>
    ///     A simple immutable storage utility to hold multiple strings.
    /// </summary>
    public sealed class StringCollection : ReadOnlyCollectionBase, IStringCollection
    {
        public static readonly IStringCollection Empty = new StringCollection();

        /// <summary>
        ///     Creates a new empty instance.
        /// </summary>
        public StringCollection()
        {
        } // StringCollection

        /// <summary>
        ///     Creates a new instance with all the items of the given collection.
        /// </summary>
        /// <param name="collection">
        ///     the items to add. may not be null. any non-string items
        ///     in this collection will be returned as null when trying to access them later.
        /// </param>
        public StringCollection(ICollection collection)
        {
            InnerList.AddRange(collection);
        } // StringCollection

        /// <summary>
        ///     Creates a new instance with all the items of the given collection.
        /// </summary>
        /// <param name="collection">
        ///     the items to add. may not be null. any non-string items
        ///     in this collection will be returned as null when trying to access them later.
        /// </param>
        public StringCollection(IStringCollection collection)
        {
            if (collection == null)
                throw new ArgumentNullException(nameof(collection));
            InnerList.Capacity = collection.Count;
            foreach (string item in collection)
                InnerList.Add(item);
        } // StringCollection

        /// <summary>
        ///     Creates a new instance with all the items of the given array.
        /// </summary>
        /// <param name="items">
        ///     the items to add. may be null or empty. any null items in
        ///     this array will be added too.
        /// </param>
        public StringCollection(params string[] items)
        {
            if (items != null)
                foreach (var item in items)
                    InnerList.Add(item);
        } // StringCollection

        /// <summary>
        ///     Index access to the items of this collection.
        /// </summary>
        /// <param name="index">the index of the item to retrieve</param>
        /// <returns>the item at the given position</returns>
        public string this[int index]
        {
            get { return InnerList[index] as string; }
        } // this[]

        public string FormatCommaSeparated()
        {
            var buf = new StringBuilder();
            var first = true;
            foreach (string item in InnerList)
            {
                if (first)
                    first = false;
                else
                    buf.Append(", ");
                buf.Append(item);
            }
            return buf.ToString();
        } // FormatCommaSeparated

        /// <summary>
        ///     Copies this collections items to the given array.
        /// </summary>
        /// <param name="array">the target array</param>
        /// <param name="index">the target index</param>
        public void CopyTo(string[] array, int index)
        {
            InnerList.CopyTo(array, index);
        } // CopyTo

        public int IndexOf(string test)
        {
            var count = InnerList.Count;
            for (var i = 0; i < count; i++)
                if (CompareTool.AreEqual(test, InnerList[i]))
                    return i;
            return -1;
        } // IndexOf

        public bool Contains(string test)
        {
            return IndexOf(test) >= 0;
        } // Contains

        /// <summary>
        ///     Adds the given item.
        /// </summary>
        /// <param name="item">the item to add</param>
        /// <returns>the insertion position</returns>
        public int Add(string item)
        {
            return InnerList.Add(item);
        } // Add

        /// <summary>
        ///     Removes the given item.
        /// </summary>
        /// <param name="item">the item to remove</param>
        public void Remove(string item)
        {
            InnerList.Remove(item);
        } // Remove

        /// <summary>
        ///     Removes the given item.
        /// </summary>
        /// <param name="index">the index of the item to remove</param>
        public void RemoveAt(int index)
        {
            InnerList.RemoveAt(index);
        } // RemoveAt

        /// <summary>
        ///     Adds the given item at the given position.
        /// </summary>
        /// <param name="item">the item to add</param>
        /// <param name="pos">the position to insert the new item into</param>
        public void Add(string item, int pos)
        {
            InnerList.Insert(pos, item);
        } // Add

        /// <summary>
        ///     Adds all items in the given list to this instance.
        /// </summary>
        /// <param name="items">the items to add</param>
        public void AddAll(IStringCollection items)
        {
            if (items == null)
                throw new ArgumentNullException(nameof(items));
            foreach (string item in items)
                InnerList.Add(item);
        } // AddAll

        /// <summary>
        ///     Adds all items in the given comma separated list to this instance.
        /// </summary>
        /// <param name="commaSeparatedList">the items to add</param>
        public void AddCommaSeparated(string commaSeparatedList)
        {
            if (!string.IsNullOrEmpty(commaSeparatedList))
            {
                var items = commaSeparatedList.Split(',');
                for (var i = 0; i < items.Length; i++)
                    InnerList.Add(items[i].Trim());
            }
        } // AddCommaSeparated

        public static StringCollection FromCommaSeparated(string commaSeparatedList)
        {
            var fromCommaSeparated = new StringCollection();
            fromCommaSeparated.AddCommaSeparated(commaSeparatedList);
            return fromCommaSeparated;
        } // FromCommaSeparated

        /// <summary>
        ///     Removes all items in the given list from this instance.
        /// </summary>
        /// <param name="items">the items to remove</param>
        public void RemoveAll(IStringCollection items)
        {
            if (items == null)
                throw new ArgumentNullException(nameof(items));
            foreach (string item in items)
                Remove(item);
        } // RemoveAll

        public void Clear()
        {
            InnerList.Clear();
        } // Clear

        public void Sort()
        {
            var count = InnerList.Count;
            var items = new string[count];
            InnerList.CopyTo(items);
            Array.Sort(items);
            for (var i = 0; i < count; i++)
                InnerList[i] = items[i];
        } // Sort

        public override bool Equals(object obj)
        {
            return CollectionTool.AreEqual(this, obj);
        } // Equals

        public override int GetHashCode()
        {
            return CollectionTool.ComputeHashCode(this);
        } // GetHashCode

        /// <summary>
        ///     Lists the contents of this collection.
        /// </summary>
        /// <returns>a string with the items of this collection</returns>
        public override string ToString()
        {
            return CollectionTool.ToString(this);
        } // ToString
    }
}