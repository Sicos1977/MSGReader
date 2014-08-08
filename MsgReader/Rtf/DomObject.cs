using System;
using System.Collections.Generic;

namespace MsgReader.Rtf
{
    /// <summary>
    /// RTF Dom object
    /// </summary>
    internal class DomObject : DomElement
    {
        #region Fields
        private Dictionary<string, string> _customAttributes = new Dictionary<string, string>();
        #endregion

        #region Properties
        /// <summary>
        /// Custom attributes
        /// </summary>
        public Dictionary<string, string> CustomAttributes
        {
            get { return _customAttributes ?? (_customAttributes = new Dictionary<string, string>()); }
            set
            {
                _customAttributes = value;
            }
        }

        /// <summary>
        /// Type
        /// </summary>
        public RtfObjectType Type { get; set; }

        /// <summary>
        /// Class name
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Content as byte array
        /// </summary>
        public byte[] Content { get; set; }

        /// <summary>
        /// Content type
        /// </summary>
        public string ContentText
        {
            get
            {
                if (Content == null || Content.Length == 0)
                    return null;

                return System.Text.Encoding.Default.GetString(Content);
            }
        }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Scaling X factor
        /// </summary>
        public int ScaleX { get; set; }

        /// <summary>
        /// Scaling Y factor
        /// </summary>
        public int ScaleY { get; set; }

        /// <summary>
        /// Result
        /// </summary>
        public ElementContainer Result
        {
            get
            {
                foreach (DomElement element in Elements)
                {
                    if (element is ElementContainer)
                    {
                        var container = (ElementContainer)element;
                        if (container.Name == Consts.Result)
                            return container;
                    }
                }
                return null;
            }
        }

        #endregion

        #region Constructor
        public DomObject()
        {
            ScaleX = 100;
            ScaleY = 100;
            Height = 0;
            Width = 0;
            Content = null;
            Name = null;
            ClassName = null;
            Type = RtfObjectType.Emb;
        }
        #endregion
        
        #region ToString
        public override string ToString()
        {
            var text = "Object:" + Width + "*" + Height;
            if (Content != null && Content.Length > 0)
                text = text + " " + Convert.ToDouble(Content.Length / 1024.0).ToString("0.00") + "KB";
            return text;
        }
        #endregion
    }
}
