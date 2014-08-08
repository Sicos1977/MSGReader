using System;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Rtf paragraph element
    /// </summary>
    internal class DomParagraph : DomElement
    {
        #region Properties
        /// <summary>
        /// If it is generated from a template
        /// </summary>
        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool IsTemplateGenerated { get; internal set; }

        /// <summary>
        /// Format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }
        #endregion

        #region Constructor
        public DomParagraph()
        {
            Format = new DocumentFormatInfo();
            IsTemplateGenerated = false;
        }
        #endregion

        #region InnerText
        public override string InnerText
        {
            get { return base.InnerText + Environment.NewLine; }
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("Paragraph");
            if (Format != null)
            {
                stringBuilder.Append("(" + Format.Align + ")");
                if (Format.ListId >= 0)
                    stringBuilder.Append("ListID:" + Format.ListId);
            }

            return stringBuilder.ToString();
        }
        #endregion
    }
}