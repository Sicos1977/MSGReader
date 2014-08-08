using System;
using System.Collections;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF raw document
    /// </summary>
    internal class RawDocument : NodeGroup
    {
        #region Fields
        // ReSharper disable MemberCanBePrivate.Global
        // ReSharper disable FieldCanBeMadeReadOnly.Global
        /// <summary>
        /// Text encoding for current associated font
        /// </summary>
        private Encoding _associatedFontChartset;

        /// <summary>
        /// Color table
        /// </summary>
        protected ColorTable InternalColorTable = new ColorTable();

        private Encoding _encoding;

        /// <summary>
        /// Text encoding for current font
        /// </summary>
        private Encoding _fontChartset;

        /// <summary>
        /// font table
        /// </summary>
        protected Table InternalFontTable = new Table();

        /// <summary>
        /// document information
        /// </summary>
        protected DocumentInfo InternalInfo = new DocumentInfo();
        // ReSharper restore FieldCanBeMadeReadOnly.Global
        #endregion

        #region Properties
        /// <summary>
        /// The owner of this document
        /// </summary>
        public override RawDocument OwnerDocument
        {
            get { return this; }
            set { }
        }

        /// <summary>
        /// No parent node
        /// </summary>
        public override NodeGroup Parent
        {
            get { return null; }
            set { }
        }

        /// <summary>
        /// color table
        /// </summary>
        public ColorTable ColorTable
        {
            get { return InternalColorTable; }
        }

        /// <summary>
        /// font table
        /// </summary>
        public Table FontTable
        {
            get { return InternalFontTable; }
        }

        /// <summary>
        /// Document information
        /// </summary>
        public DocumentInfo Info
        {
            get { return InternalInfo; }
        }

        /// <summary>
        /// Text encoding
        /// </summary>
        public Encoding Encoding
        {
            get
            {
                if (_encoding == null)
                {
                    var node = InternalNodes[Consts.Ansicpg];
                    if (node != null && node.HasParameter)
                        _encoding = Encoding.GetEncoding(node.Parameter);
                }

                return _encoding ?? (_encoding = Encoding.Default);
            }
        }

        /// <summary>
        /// Current text encoding
        /// </summary>
        internal Encoding RuntimeEncoding
        {
            get
            {
                if (_fontChartset != null)
                    return _fontChartset;

                return _associatedFontChartset ?? Encoding;
            }
        }
        // ReSharper restore MemberCanBePrivate.Global
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize instance
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RawDocument()
        {
            InternalOwnerDocument = this;
            // ReSharper disable once DoNotCallOverridableMethodsInConstructor
            Parent = null;
            InternalColorTable.CheckValueExistWhenAdd = false;
        }
        #endregion

        #region ReadFontTable
        /// <summary>
        /// Read font table
        /// </summary>
        /// <param name="group"></param>
        private void ReadFontTable(NodeGroup group)
        {
            InternalFontTable.Clear();
            foreach (Node node in group.Nodes)
            {
                if (node is NodeGroup)
                {
                    var index = -1;
                    string name = null;
                    var charset = 0;

                    foreach (Node item in node.Nodes)
                    {
                        if (item.Keyword == "f" && item.HasParameter)
                            index = item.Parameter;
                        else if (item.Keyword == Consts.Fcharset)
                            charset = item.Parameter;
                        else if (item.Type == RtfNodeType.Text)
                        {
                            if (!string.IsNullOrEmpty(item.Keyword))
                            {
                                name = item.Keyword;
                                break;
                            }
                        }
                    }

                    if (index >= 0 && name != null)
                    {
                        if (name.EndsWith(";"))
                            name = name.Substring(0, name.Length - 1);

                        name = name.Trim();
                        //System.Console.WriteLine( "Index:" + index + "  Name:" + name );
                        var font = new Font(index, name) {Charset = charset};
                        InternalFontTable.Add(font);
                    }
                }
            }
        }
        #endregion

        #region ReadColorTable
        /// <summary>
        /// Read color table
        /// </summary>
        /// <param name="group"></param>
        private void ReadColorTable(NodeGroup group)
        {
            InternalColorTable.Clear();
            var r = -1;
            var g = -1;
            var b = -1;

            foreach (Node node in group.Nodes)
            {
                if (node.Keyword == "red")
                    r = node.Parameter;
                else if (node.Keyword == "green")
                    g = node.Parameter;
                else if (node.Keyword == "blue")
                    b = node.Parameter;
            
                if (node.Keyword == ";")
                {
                    if (r >= 0 && g >= 0 && b >= 0)
                    {
                        var c = Color.FromArgb(255, r, g, b);
                        InternalColorTable.Add(c);
                        r = -1;
                        g = -1;
                        b = -1;
                    }
                }
            }

            if (r >= 0 && g >= 0 && b >= 0)
            {
                // Read the last color
                var c = Color.FromArgb(255, r, g, b);
                InternalColorTable.Add(c);
            }
        }
        #endregion

        #region ReadDocumentInfo
        /// <summary>
        /// Read document information
        /// </summary>
        /// <param name="group"></param>
        private void ReadDocumentInfo(NodeGroup group)
        {
            InternalInfo.Clear();
            //var nodeList = group.GetAllNodes(false);
            
            foreach (Node node in group.Nodes)
            {
                if ((node is NodeGroup) == false)
                    continue;
                if (node.Keyword == "creatim")
                    InternalInfo.CreationTime = ReadDateTime(node);
                else if (node.Keyword == "revtim")
                    InternalInfo.RevisionTime = ReadDateTime(node);
                else if (node.Keyword == "printim")
                    InternalInfo.PrintTime = ReadDateTime(node);
                else if (node.Keyword == "buptim")
                    InternalInfo.BackupTime = ReadDateTime(node);
                else
                    InternalInfo.SetInfo(node.Keyword, node.HasParameter ? node.Parameter.ToString(CultureInfo.InvariantCulture) : node.Nodes.Text);
            }
        }
        #endregion

        #region ReadDateTime
        /// <summary>
        /// Read date and time from node
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        private DateTime ReadDateTime(Node node)
        {
            var yr = node.Nodes.GetParameter("yr", 1900);
            var mo = node.Nodes.GetParameter("mo", 1);
            var dy = node.Nodes.GetParameter("dy", 1);
            var hr = node.Nodes.GetParameter("hr", 0);
            var min = node.Nodes.GetParameter("min", 0);
            var sec = node.Nodes.GetParameter("sec", 0);
            return new DateTime(yr, mo, dy, hr, min, sec);
        }
        #endregion

        #region Load
        /// <summary>
        /// Load rtf from string
        /// </summary>
        /// <param name="text">Text in rtf format</param>
        // ReSharper disable once UnusedMember.Global
        public void LoadRTFText(string text)
        {
            _encoding = null;
            using (var reader = new Reader())
            {
                if (reader.LoadRTFText(text))
                    Load(reader);
            }
        }

        /// <summary>
        /// Load rtf from file
        /// </summary>
        /// <param name="fileName">file name</param>
        // ReSharper disable once UnusedMember.Global
        public void Load(string fileName)
        {
            _encoding = null;
            using (var reader = new Reader())
            {
                if (reader.LoadRTFFile(fileName))
                    Load(reader);
            }
        }

        /// <summary>
        /// Load rtf from text reader
        /// </summary>
        /// <param name="reader"></param>
        public void Load(TextReader reader)
        {
            var rdr = new Reader();
            rdr.LoadReader(reader);
            Load(rdr);
        }

        /// <summary>
        /// Load rtf from rtf reader
        /// </summary>
        /// <param name="reader">RTF text reader</param>
        // ReSharper disable once MemberCanBePrivate.Global
        public void Load(Reader reader)
        {
            InternalNodes.Clear();
            var groups = new Stack();
            NodeGroup newGroup = null;
            while (reader.ReadToken() != null)
            {
                if (reader.TokenType == RtfTokenType.GroupStart)
                {
                    // Begin group
                    if (newGroup == null)
                        newGroup = this;
                    else
                    {
                        newGroup = new NodeGroup {OwnerDocument = this};
                    }
                    if (newGroup != this)
                    {
                        var group = (NodeGroup) groups.Peek();
                        group.AppendChild(newGroup);
                    }
                    groups.Push(newGroup);
                }
                else if (reader.TokenType == RtfTokenType.GroupEnd)
                {
                    // end group
                    newGroup = (NodeGroup) groups.Pop();
                    newGroup.MergeText();
                    // ReSharper disable once CSharpWarnings::CS0183
                    if (newGroup.FirstNode is Node)
                    {
                        switch (newGroup.Keyword)
                        {
                            case Consts.Fonttbl:
                                // Read font table
                                ReadFontTable(newGroup);
                                break;

                            case Consts.Colortbl:
                                // Read color table
                                ReadColorTable(newGroup);
                                break;

                            case Consts.Info:
                                // Read document information
                                ReadDocumentInfo(newGroup);
                                break;
                        }
                    }

                    if (groups.Count > 0)
                        newGroup = (NodeGroup) groups.Peek();
                    else
                        break;
                    //NewGroup.MergeText();
                }
                else
                {
                    // Read content
                    var newNode = new Node(reader.CurrentToken) {OwnerDocument = this};
                    
                    if (newGroup != null) newGroup.AppendChild(newNode);

                    switch (newNode.Keyword)
                    {
                        case Consts.F:
                        {
                            var font = FontTable[newNode.Parameter];
                            _fontChartset = font != null ? font.Encoding : null;
                            //myFontChartset = RTFFont.GetRTFEncoding( NewNode.Parameter );
                        }
                            break;

                        case Consts.Af:
                        {
                            var font = FontTable[newNode.Parameter];
                            _associatedFontChartset = font != null ? font.Encoding : null;
                        }
                            break;

                    }

                }
            } 

            while (groups.Count > 0)
            {
                newGroup = (NodeGroup) groups.Pop();
                newGroup.MergeText();
            }
            //this.UpdateInformation();
        }
        #endregion

        #region Write
        /// <summary>
        /// Write rtf
        /// </summary>
        /// <param name="writer">RTF writer</param>
        public override void Write(Writer writer)
        {
            writer.Encoding = Encoding;
            base.Write(writer);
        }
        #endregion

        #region Save
        /// <summary>
        /// Save rtf to a file
        /// </summary>
        /// <param name="fileName">file name</param>
        public void Save(string fileName)
        {
            using (var writer = new Writer(fileName))
                Write(writer);
        }

        /// <summary>
        /// Save rtf to a stream
        /// </summary>
        /// <param name="stream">stream</param>
        public void Save(Stream stream)
        {
            using (var writer = new Writer(new StreamWriter(stream, Encoding)))
                Write(writer);
        }
        #endregion
    }
}