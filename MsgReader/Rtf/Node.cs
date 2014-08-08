
namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF parser node
    /// </summary>
    internal class Node
    {
        #region Fields
        protected RawDocument InternalOwnerDocument = null;
        protected NodeGroup InternalParent = null;
        protected string InternalKeyword = null;
        #endregion

        #region Properties
        /// <summary>
        /// parent node
        /// </summary>
        public virtual NodeGroup Parent
        {
            get { return InternalParent; }
            set { InternalParent = value; }
        }

        /// <summary>
        /// raw document which owner this node
        /// </summary>
        public virtual RawDocument OwnerDocument
        {
            get { return InternalOwnerDocument; }
            set
            {
                InternalOwnerDocument = value;
                if (Nodes != null)
                {
                    foreach (Node node in Nodes)
                    {
                        node.OwnerDocument = value;
                    }
                }
            }
        }

        /// <summary>
        /// Keyword
        /// </summary>
        public virtual string Keyword
        {
            get { return InternalKeyword; }
            set { InternalKeyword = value; }
        }

        /// <summary>
        /// Whether this node has parameters
        /// </summary>
        public virtual bool HasParameter { get; set; }

        /// <summary>
        /// Paramter value
        /// </summary>
        public virtual int Parameter { get; protected set; }

        /// <summary>
        /// Child nodes
        /// </summary>
        public virtual NodeList Nodes
        {
            get { return null; }
        }


        /// <summary>
        /// Index
        /// </summary>
        public int Index
        {
            get
            {
                return InternalParent == null ? 0 : InternalParent.Nodes.IndexOf(this);
            }
        }

        /// <summary>
        /// node type
        /// </summary>
        public RtfNodeType Type { get; protected set; }

        /// <summary>
        /// Previous node in parent nodes list
        /// </summary>
        public Node PreviousNode
        {
            get
            {
                if (InternalParent != null)
                {
                    var index = InternalParent.Nodes.IndexOf(this);
                    if (index > 0)
                    {
                        return InternalParent.Nodes[index - 1];
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Next node in parent nodes list
        /// </summary>
        public Node NextNode
        {
            get
            {
                if (InternalParent != null)
                {
                    var index = InternalParent.Nodes.IndexOf(this);
                    if (index >= 0 && index < InternalParent.Nodes.Count - 1)
                        return InternalParent.Nodes[index + 1];
                }
                return null;
            }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// initialize instance
        /// </summary>
        public Node()
        {
            Type = RtfNodeType.None;
        }

        public Node(RtfNodeType type, string key)
        {
            Type = type;
            InternalKeyword = key;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Node(Token token)
        {
            InternalKeyword = token.Key;
            // ReSharper disable DoNotCallOverridableMethodsInConstructor
            HasParameter = token.HasParam;
            Parameter = token.Param;
            // ReSharper restore DoNotCallOverridableMethodsInConstructor
            switch (token.Type)
            {
                case RtfTokenType.Control:
                    Type = RtfNodeType.Control;
                    break;
                case RtfTokenType.Keyword:
                    Type = RtfNodeType.Keyword;
                    break;
                case RtfTokenType.ExtKeyword:
                    Type = RtfNodeType.ExtKeyword;
                    break;
                case RtfTokenType.Text:
                    Type = RtfNodeType.Text;
                    break;
                default:
                    Type = RtfNodeType.Text;
                    break;
            }
        }
        #endregion

        #region Write
        /// <summary>
        /// Write to rtf document
        /// </summary>
        /// <param name="writer">RTF text writer</param>
        public virtual void Write(Writer writer)
        {
            switch (Type)
            {
                case RtfNodeType.ExtKeyword:
                case RtfNodeType.Keyword:
                case RtfNodeType.Control:

                    if (HasParameter)
                        writer.WriteKeyword(InternalKeyword + Parameter, Type == RtfNodeType.ExtKeyword);
                    else
                        writer.WriteKeyword(InternalKeyword, Type == RtfNodeType.ExtKeyword);
                    break;
                
                case RtfNodeType.Text:
                    writer.WriteText(InternalKeyword);
                    break;
            }
        }
        #endregion
    }
}