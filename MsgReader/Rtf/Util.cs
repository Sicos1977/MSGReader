using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace DocumentServices.Modules.Readers.MsgReader.Rtf
{
    /// <summary>
    /// Some utility functions
    /// </summary>
    internal static class Util
    {
        #region HasContentElement
        /// <summary>
        /// Checks if the root element has content elemens
        /// </summary>
        /// <param name="rootElement"></param>
        /// <returns>True when there are content elements</returns>
        public static bool HasContentElement(DomElement rootElement)
        {
            if (rootElement.Elements.Count == 0)
            {
                return false;
            }
            if (rootElement.Elements.Count == 1)
            {
                if (rootElement.Elements[0] is DomParagraph)
                {
                    var p = (DomParagraph) rootElement.Elements[0];
                    if (p.Elements.Count == 0)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion

        #region GetRtfImage
        /// <summary>
        /// Use the EmfToWmfBits function in the GDI+ specification to convert a 
        /// Enhanced Metafile to a Windows Metafile
        /// </summary>
        /// <param name="hEmf">
        /// A handle to the Enhanced Metafile to be converted
        /// </param>
        /// <param name="bufferSize">
        /// The size of the buffer used to store the Windows Metafile bits returned
        /// </param>
        /// <param name="buffer">
        /// An array of bytes used to hold the Windows Metafile bits returned
        /// </param>
        /// <param name="mappingMode">
        /// The mapping mode of the image.  This control uses MM_ANISOTROPIC.
        /// </param>
        /// <param name="flags">
        /// Flags used to specify the format of the Windows Metafile returned
        /// </param>
        [DllImport("gdiplus.dll")]
        private static extern uint GdipEmfToWmfBits(IntPtr hEmf, uint bufferSize, byte[] buffer, int mappingMode, EmfToWmfBitsFlags flags);
        // Specifies the flags/options for the unmanaged call to the GDI+ method
        // Metafile.EmfToWmfBits().


        /// <summary>
        /// Wraps the image in an Enhanced Metafile by drawing the image onto the
        /// graphics context, then converts the Enhanced Metafile to a Windows
        /// Metafile, and finally appends the bits of the Windows Metafile in HEX
        /// to a string and returns the string.
        /// </summary>
        /// <param name="image"></param>
        /// <returns>
        /// A string containing the bits of a Windows Metafile in HEX
        /// </returns>
        // ReSharper disable UnusedMember.Local
        public static string GetRtfImage(Image image)
        // ReSharper restore UnusedMember.Local
        {
            // Ensures that the metafile maintains a 1:1 aspect ratio
            //const int MM_ISOTROPIC = 7;

            // Allows the x-coordinates and y-coordinates of the metafile to be adjusted
            // independently
            const int mmAnisotropic = 8;

            // Used to store the enhanced metafile
            MemoryStream stream = null;

            // Used to create the metafile and draw the image
            Graphics graphics = null;

            // The enhanced metafile
            Metafile metaFile = null;

            // Handle to the device context used to create the metafile

            try
            {
                var rtf = new StringBuilder();
                stream = new MemoryStream();

                // Get a graphics context from the RichTextBox
                using (graphics = Graphics.FromHwnd(new IntPtr(0)))
                {
                    // Get the device context from the graphics context
                    var hdc = graphics.GetHdc();

                    // Create a new Enhanced Metafile from the device context
                    metaFile = new Metafile(stream, hdc);

                    // Release the device context
                    graphics.ReleaseHdc(hdc);
                }

                // Get a graphics context from the Enhanced Metafile
                using (graphics = Graphics.FromImage(metaFile))
                {
                    // Draw the image on the Enhanced Metafile
                    graphics.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height));
                }

                // Get the handle of the Enhanced Metafile
                var hEmf = metaFile.GetHenhmetafile();

                // A call to EmfToWmfBits with a null buffer return the size of the
                // buffer need to store the WMF bits.  Use this to get the buffer
                // size.
                var bufferSize = GdipEmfToWmfBits(hEmf, 0, null, mmAnisotropic,
                    EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);

                // Create an array to hold the bits
                var buffer = new byte[bufferSize];

                // A call to EmfToWmfBits with a valid buffer copies the bits into the
                // buffer an returns the number of bits in the WMF.  
                GdipEmfToWmfBits(hEmf, bufferSize, buffer, mmAnisotropic,EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);

                // Append the bits to the RTF string
                foreach (byte t in buffer)
                    rtf.Append(String.Format("{0:X2}", t));

                return rtf.ToString();
            }
            finally
            {
                if (graphics != null)
                    graphics.Dispose();

                if (metaFile != null)
                    metaFile.Dispose();

                if (stream != null)
                    stream.Close();
            }
        }
        #endregion

        #region Enum EmfToWmfBitsFlags
        private enum EmfToWmfBitsFlags
        {
            // ReSharper disable UnusedMember.Local
            // Use the default conversion
            EmfToWmfBitsFlagsDefault = 0x00000000,

            // Embedded the source of the EMF metafiel within the resulting WMF
            // metafile
            EmfToWmfBitsFlagsEmbedEmf = 0x00000001,

            // Place a 22-byte header in the resulting WMF file.  The header is
            // required for the metafile to be considered placeable.
            EmfToWmfBitsFlagsIncludePlaceable = 0x00000002,

            // Don't simulate clipping by using the XOR operator.
            EmfToWmfBitsFlagsNoXorClip = 0x00000004
            // ReSharper restore UnusedMember.Local
        };
        #endregion
    }
}