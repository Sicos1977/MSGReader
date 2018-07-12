namespace MsgReader.Outlook
{
    /// <summary>
    /// Used as a temporary placeholder for information about the inline attachents
    /// </summary>
    internal sealed class InlineAttachment
    {
        #region Properties
        /// <summary>
        /// Returns the rendering position for the attachmnt
        /// </summary>
        public int? RenderingPosition { get; private set; }

        /// <summary>
        /// Returns the name of the icon when this attachment is part of an RTF body and is
        /// shown as an icon
        /// </summary>
        public string IconFileName { get; private set; }

        /// <summary>
        /// Returns the name for the attachment
        /// </summary>
        public string AttachmentFileName { get; private set; }

        /// <summary>
        /// Returns the full name for the attachment
        /// </summary>
        public string FullName { get; private set; }
        #endregion

        #region Constructors
        public InlineAttachment(int renderingPosition,
            string attachmentFileName)
        {
            RenderingPosition = renderingPosition;
            AttachmentFileName = attachmentFileName;
            FullName = attachmentFileName;
        }

        public InlineAttachment(string iconFileName,
            string attachmentFileName,
            string fullName)
        {
            IconFileName = iconFileName;
            AttachmentFileName = attachmentFileName;
            FullName = fullName;
        }
        #endregion
    }
}
