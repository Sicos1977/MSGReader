using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace MsgReader.Rtf
{
    /// <summary>
    /// Some utility functions
    /// </summary>
    internal static class Util
    {
        #region Private static class NativeMethofs
        private static class NativeMethods
        {
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
            public static extern uint GdipEmfToWmfBits(IntPtr hEmf, uint bufferSize, byte[] buffer, int mappingMode, EmfToWmfBitsFlags flags);
            // Specifies the flags/options for the unmanaged call to the GDI+ method
            // Metafile.EmfToWmfBits().            
        }
        #endregion

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
        {
            // Allows the x-coordinates and y-coordinates of the metafile to be adjusted
            // independently
            const int mmAnisotropic = 8;

            var rtf = new StringBuilder();
            var stream = new MemoryStream();

            // Get a graphics context from the RichTextBox
            Graphics graphics;
            Metafile metaFile;
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
                graphics.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height));

            // Get the handle of the Enhanced Metafile
            var hEmf = metaFile.GetHenhmetafile();

            // A call to EmfToWmfBits with a null buffer return the size of the
            // buffer need to store the WMF bits.  Use this to get the buffer
            // size.
            var bufferSize = NativeMethods.GdipEmfToWmfBits(hEmf, 0, null, mmAnisotropic,
                EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);

            // Create an array to hold the bits
            var buffer = new byte[bufferSize];

            // A call to EmfToWmfBits with a valid buffer copies the bits into the
            // buffer an returns the number of bits in the WMF.  
            NativeMethods.GdipEmfToWmfBits(hEmf, bufferSize, buffer, mmAnisotropic,
                EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);

            // Append the bits to the RTF string
            foreach (var b in buffer)
                rtf.Append(String.Format("{0:X2}", b));

            return rtf.ToString();
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