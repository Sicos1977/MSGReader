using System;
using System.Collections;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF node list
    /// </summary>
    internal class NodeList : CollectionBase
    {
        #region Properties
        /// <summary>
        /// Get node by index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public Node this[int index]
        {
            get { return (Node) List[index]; }
        }

        /// <summary>
        /// Get node by keyword
        /// </summary>
        /// <param name="keyWord"></param>
        /// <returns></returns>
        public Node this[string keyWord]
        {
            get
            {
                foreach (Node node in this)
                    if (node.Keyword == keyWord)
                        return node;

                return null;
            }
        }

        /// <summary>
        /// Get node by type
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public Node this[Type type]
        {
            get
            {
                foreach (Node node in this)
                {
                    if (type == node.GetType())
                        return node;
                }
                return null;
            }
        }

        /// <summary>
        /// Get the node text
        /// </summary>
        public string Text
        {
            get
            {
                var stringBuilder = new StringBuilder();
                foreach (Node node in this)
                {
                    if (node.Type == RtfNodeType.Text)
                        stringBuilder.Append(node.Keyword);
                    else if (node is NodeGroup)
                    {
                        var txt = node.Nodes.Text;
                        if (txt != null)
                            stringBuilder.Append(txt);
                    }
                }
                return stringBuilder.ToString();
            }
        }

        /// <summary>
        /// Get node's parameter value special keyword, if it is not found , return default value.
        /// </summary>
        /// <param name="key">Keyword</param>
        /// <param name="defaultValue">Default value</param>
        /// <returns>Parameter value</returns>
        public int GetParameter(string key, int defaultValue)
        {
            foreach (Node node in this)
                if (node.Keyword == key && node.HasParameter)
                    return node.Parameter;

            return defaultValue;
        }

        /// <summary>
        /// Detect whether the node exist in this list
        /// </summary>
        /// <param name="key">keyword</param>
        /// <returns>True or false</returns>
        public bool ContainsKey(string key)
        {
            return this[key] != null;
        }

        /// <summary>
        /// Get index of node in this list
        /// </summary>
        /// <param name="node">node</param>
        /// <returns>index , if node does no in this list , return -1</returns>
        public int IndexOf(Node node)
        {
            return List.IndexOf(node);
        }

        /// <summary>
        /// Add node list range
        /// </summary>
        /// <param name="nodeList">A node list</param>
        internal void AddRange(NodeList nodeList)
        {
            InnerList.AddRange(nodeList);
        }

        /// <summary>
        /// Add node
        /// </summary>
        /// <param name="node">Node</param>
        internal void Add(Node node)
        {
            //node.OwnerList = this ;
            List.Add(node);
        }

        /// <summary>
        /// Remove node
        /// </summary>
        /// <param name="node">Node</param>
        internal void Remove(Node node)
        {
            List.Remove(node);
        }

        /// <summary>
        /// Insert node
        /// </summary>
        /// <param name="index">Index</param>
        /// <param name="node">Node</param>
        internal void Insert(int index, Node node)
        {
            List.Insert(index, node);
        }
        #endregion
    }
}