using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf text element
    /// </summary>
    internal class DomText : DomElement
    {
        #region Properties
        /// <summary>
        /// Format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }

        /// <summary>
        /// Text
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// Inner text
        /// </summary>
        public override string InnerText
        {
            get { return Text; }
        }
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize instance
        /// </summary>
        public DomText()
        {
            Format = new DocumentFormatInfo();
            // Text element can not contains any child element
            Locked = true;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("Text");
            if (Format != null)
            {
                if (Format.Hidden)
                    stringBuilder.Append("(Hidden)");
            }
            stringBuilder.Append(":" + Text);
            return stringBuilder.ToString();
        }
        #endregion
    }
}