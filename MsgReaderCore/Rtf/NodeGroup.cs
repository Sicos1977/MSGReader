//
// NodeGroup.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF node group
    /// </summary>
    internal class NodeGroup : Node
    {
        #region Fields
        /// <summary>
        /// child node list
        /// </summary>
        protected NodeList InternalNodes = new NodeList();
        #endregion

        #region Constructor
        /// <summary>
        /// initialize instance
        /// </summary>
        public NodeGroup()
        {
            Type = RtfNodeType.Group;
        }
        #endregion

        #region Properties
        /// <summary>
        /// child node list
        /// </summary>
        public override NodeList Nodes => InternalNodes;

        /// <summary>
        /// first child node
        /// </summary>
        public Node FirstNode => InternalNodes.Count > 0 ? InternalNodes[0] : null;

        public override string Keyword => InternalNodes.Count > 0 ? InternalNodes[0].Keyword : null;

        public override bool HasParameter => InternalNodes.Count > 0 && InternalNodes[0].HasParameter;

        public override int Parameter => InternalNodes.Count > 0 ? InternalNodes[0].Parameter : 0;


        public virtual string Text
        {
            get
            {
                var stringBuilder = new StringBuilder();
                foreach (Node node in InternalNodes)
                {
                    if (node is NodeGroup)
                        stringBuilder.Append(((NodeGroup) node).Text);
                    if (node.Type == RtfNodeType.Text)
                        stringBuilder.Append(node.Keyword);
                }
                return stringBuilder.ToString();
            }
        }
        #endregion

        #region GetAllNodes
        /// <summary>
        /// Get all child node deeply
        /// </summary>
        /// <param name="includeGroupNode">contains group type node</param>
        /// <returns>child nodes list</returns>
        public NodeList GetAllNodes(bool includeGroupNode)
        {
            var nodeList = new NodeList();
            AddAllNodes(nodeList, includeGroupNode);
            return nodeList;
        }
        #endregion

        #region AddAllNodes
        private void AddAllNodes(NodeList nodeList, bool includeGroupNode)
        {
            foreach (Node node in InternalNodes)
            {
                if (node is NodeGroup)
                {
                    if (includeGroupNode)
                        nodeList.Add(node);
                    ((NodeGroup) node).AddAllNodes(nodeList, includeGroupNode);
                }
                else
                    nodeList.Add(node);
            }
        }
        #endregion

        #region MergeText
        internal void MergeText()
        {
            var nodeList = new NodeList();
            var stringBuilder = new StringBuilder();
            var byteBuffer = new ByteBuffer();

            foreach (Node node in InternalNodes)
            {
                if (node.Type == RtfNodeType.Text)
                {
                    AddString(stringBuilder, byteBuffer);
                    stringBuilder.Append(node.Keyword);
                    continue;
                }
                
                if (node.Type == RtfNodeType.Control
                    && node.Keyword == "\'"
                    && node.HasParameter)
                {
                    byteBuffer.Add((byte) node.Parameter);
                    continue;
                }
              
                if (node.Type == RtfNodeType.Control || node.Type == RtfNodeType.Keyword)
                {
                    if (node.Keyword == "tab")
                    {
                        AddString(stringBuilder, byteBuffer);
                        stringBuilder.Append('\t');
                        continue;
                    }
                    if (node.Keyword == "emdash")
                    {
                        AddString(stringBuilder, byteBuffer);
                        stringBuilder.Append('-');
                        continue;
                    }
                    if (node.Keyword == string.Empty)
                    {
                        AddString(stringBuilder, byteBuffer);
                        stringBuilder.Append('-');
                        continue;
                    }
                }
                
                AddString(stringBuilder, byteBuffer);
                
                if (stringBuilder.Length > 0)
                {
                    nodeList.Add(new Node(RtfNodeType.Text, stringBuilder.ToString()));
                    stringBuilder = new StringBuilder();
                }
                nodeList.Add(node);
            } 

            AddString(stringBuilder, byteBuffer);
            
            if (stringBuilder.Length > 0)
                nodeList.Add(new Node(RtfNodeType.Text, stringBuilder.ToString()));

            InternalNodes.Clear();
            
            foreach (Node node in nodeList)
            {
                node.Parent = this;
                node.OwnerDocument = InternalOwnerDocument;
                InternalNodes.Add(node);
            }
        }
        #endregion

        #region AddString
        private void AddString(StringBuilder stringBuilder, ByteBuffer buffer)
        {
            if (buffer.Count > 0)
            {
                var text = buffer.GetString(InternalOwnerDocument.RuntimeEncoding);
                stringBuilder.Append(text);
                buffer.Reset();
            }
        }
        #endregion

        #region Write
        /// <summary>
        /// Write content to rtf document
        /// </summary>
        /// <param name="writer">RTF text writer</param>
        public override void Write(Writer writer)
        {
            writer.WriteStartGroup();
            
            foreach (Node node in InternalNodes)
                node.Write(writer);

            writer.WriteEndGroup();
        }
        #endregion

        #region SearchKey
        /// <summary>
        /// Search child node with special keyword
        /// </summary>
        /// <param name="key">Special keyword</param>
        /// <param name="deeply">Whether search deeply</param>
        /// <returns>Found node</returns>
        public Node SearchKey(string key, bool deeply)
        {
            foreach (Node node in InternalNodes)
            {
                if (node.Type == RtfNodeType.Keyword
                    || node.Type == RtfNodeType.ExtKeyword
                    || node.Type == RtfNodeType.Control)
                {
                    if (node.Keyword == key)
                        return node;
                }
                
                if (deeply)
                {
                    if (node is NodeGroup)
                    {
                        var groupNode = (NodeGroup) node;
                        var n = groupNode.SearchKey(key, true);
                        if (n != null)
                            return n;
                    }
                }
            }
            return null;
        }
        #endregion

        #region AppendChild
        /// <summary>
        /// Append child node
        /// </summary>
        /// <param name="node">node</param>
        public void AppendChild(Node node)
        {
            CheckNodes();
            if (node == null)
                throw new ArgumentNullException(nameof(node));
            if (node == this)
                throw new ArgumentException("node != this");
            node.Parent = this;
            node.OwnerDocument = InternalOwnerDocument;
            Nodes.Add(node);
        }
        #endregion

        #region RemoveChild
        /// <summary>
        /// Remove child
        /// </summary>
        /// <param name="node">node</param>
        public void RemoveChild(Node node)
        {
            CheckNodes();
            if (node == null)
                throw new ArgumentNullException(nameof(node));
            if (node == this)
                throw new ArgumentException("node != this");
            Nodes.Remove(node);
        }
        #endregion

        #region InsertNode
        /// <summary>
        /// Insert node
        /// </summary>
        /// <param name="index">index</param>
        /// <param name="node">node</param>
        public void InsertNode(int index, Node node)
        {
            CheckNodes();
            if (node == null)
                throw new ArgumentNullException(nameof(node));
            if (node == this)
                throw new ArgumentException("node != this");
            
            node.Parent = this;
            node.OwnerDocument = InternalOwnerDocument;
            Nodes.Insert(index, node);
        }
        #endregion

        #region CheckNodes
        private void CheckNodes()
        {
            if (Nodes == null)
                throw new Exception("Child node is invalid");
        }
        #endregion
    }
}