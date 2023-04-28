using System;
using System.Net.Mime;
using System.Text;
using MsgReader.Mime.Header;
using MsgReader.Tnef.Enums;

namespace MsgReader.Tnef
{
    internal class Attachment
    {
        #region Properties
        /// <summary>
        ///     <see cref="AttachmentType"/>
        /// </summary>
        public AttachmentType Type { get; set; }

        public bool IsTextBody { get; set; }

        /// <summary>
        ///     Returns the encoding
        /// </summary>
        public Encoding Encoding { get; set; }

        /// <summary>
        ///     The creation date of the attachment or <c>null</c> when not known
        /// </summary>
        public DateTime? CreationDate { get; set; }

        /// <summary>
        ///     The modification date of the attachment or <c>null</c> when not known
        /// </summary>
        public DateTime? ModificationDate { get; set; }
        
        /// <summary>
        ///     The file name of the attachment
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        ///     The content id or <c>null</c> when not known
        /// </summary>
        public string ContentId { get; set; }

        /// <summary>
        ///     The content location
        /// </summary>
        public Uri ContentLocation { get; set; }

        /// <summary>
        ///     The content base
        /// </summary>
        public Uri ContentBase { get; set; }

        /// <summary>
        ///     <see cref="ContentDisposition"/>
        /// </summary>
        public ContentDisposition ContentDisposition { get; set; }

        /// <summary>
        ///     <see cref="ContentTransferEncoding"/>
        /// </summary>
        public ContentType ContentType { get; set; }

        /// <summary>
        ///     <see cref="ContentTransferEncoding"/>
        /// </summary>
        public ContentTransferEncoding ContentTransferEncoding { get; set; }

        /// <summary>
        ///     The attachment itself as an byte array
        /// </summary>
        public byte[] Body { get; set; }
        #endregion
    }
}
