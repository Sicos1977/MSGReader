using System;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Image element
    /// </summary>
    internal class DomImage : DomElement
    {
        #region Properties
        /// <summary>
        /// Id
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Data
        /// </summary>
        public byte[] Data { get; set; }

        /// <summary>
        /// Data as base64 format
        /// </summary>
        public string Base64Data
        {
            get
            {
                return Data == null ? null : Convert.ToBase64String(Data);
            }
            set
            {
                Data = !string.IsNullOrEmpty(value) ? Convert.FromBase64String(value) : null;
            }
        }

        /// <summary>
        /// Scale rate at the X coordinate, in percent unit.
        /// </summary>
        public int ScaleX { get; set; }

        /// <summary>
        /// Scale rate at the Y coordinate , in percent unit.
        /// </summary>
        public int ScaleY { get; set; }

        /// <summary>
        /// Desired width
        /// </summary>
        public int DesiredWidth { get; set; }

        /// <summary>
        /// Desired height
        /// </summary>
        public int DesiredHeight { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Picture type
        /// </summary>
        public RtfPictureType PicType { get; set; }

        /// <summary>
        /// format
        /// </summary>
        public DocumentFormatInfo Format { get; set; }
        #endregion

        #region Constructor
        public DomImage()
        {
            Format = new DocumentFormatInfo();
            ScaleY = 100;
            ScaleX = 100;
            PicType = RtfPictureType.Jpegblip;
            Height = 0;
            Width = 0;
            DesiredHeight = 0;
            DesiredWidth = 0;
            Data = null;
            Id = null;
        }
        #endregion

        #region ToString
        public override string ToString()
        {
            var text = "Image:" + Width + "*" + Height;
            if (Data != null && Data.Length > 0)
                text = text + " " + Convert.ToDouble( Data.Length / 1024.0).ToString("0.00") + "KB";
            return text;
        }
        #endregion
    }
}
