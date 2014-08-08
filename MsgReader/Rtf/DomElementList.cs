using System.Collections;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF element's list
    /// </summary>
    internal class DomElementList : CollectionBase
    {
        #region Properties
        /// <summary>
        /// Get the last element in the list
        /// </summary>
        public DomElement LastElement
        {
            get { return Count > 0 ? (DomElement)List[Count - 1] : null; }
        }

        /// <summary>
        /// Get the element at special index
        /// </summary>
        /// <param name="index">index</param>
        /// <returns>element</returns>
        public DomElement this[int index]
        {
            get { return (DomElement) List[index]; }
        }
        #endregion

        #region Add
        /// <summary>
        /// Add element
        /// </summary>
        /// <param name="element">element</param>
        /// <returns>index</returns>
        public int Add(DomElement element)
        {
            return List.Add(element);
        }
        #endregion

        #region Insert
        /// <summary>
        /// Insert element
        /// </summary>
        /// <param name="index">special index</param>
        /// <param name="element">element</param>
        public void Insert(int index, DomElement element)
        {
            List.Insert(index, element);
        }
        #endregion

        #region IndexOf
        /// <summary>
        /// Get the index of special element that starts with 0.
        /// </summary>
        /// <param name="element">element</param>
        /// <returns>index , if not find element , then return -1</returns>
        public int IndexOf(DomElement element)
        {
            return List.IndexOf(element);
        }
        #endregion

        #region Remove
        /// <summary>
        /// Remove element
        /// </summary>
        /// <param name="node">element</param>
        public void Remove(DomElement node)
        {
            List.Remove(node);
        }
        #endregion

        #region ToArray
        /// <summary>
        /// Return element array
        /// </summary>
        /// <returns>array</returns>
        public DomElement[] ToArray()
        {
            return (DomElement[]) InnerList.ToArray(typeof (DomElement));
        }
        #endregion
    }
}