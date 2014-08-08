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
        public override NodeList Nodes
        {
            get { return InternalNodes; }
        }

        /// <summary>
        /// first child node
        /// </summary>
        public Node FirstNode
        {
            get
            {
                return InternalNodes.Count > 0 ? InternalNodes[0] : null;
            }
        }

        public override string Keyword
        {
            get
            {
                return InternalNodes.Count > 0 ? InternalNodes[0].Keyword : null;
            }
        }

        public override bool HasParameter
        {
            get
            {
                return InternalNodes.Count > 0 && InternalNodes[0].HasParameter;
            }
        }

        public override int Parameter
        {
            get
            {
                return InternalNodes.Count > 0 ? InternalNodes[0].Parameter : 0;
            }
        }


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
                throw new ArgumentNullException("node");
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
                throw new ArgumentNullException("node");
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
                throw new ArgumentNullException("node");
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