using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal class RtfFileIcon
    {
        #region Consts
        /// <summary>
        /// Allows the x-coordinates and y-coordinates of the metafile to be adjusted independently
        /// </summary>
        private const int MmAnisotropic = 8;

        /// <summary>
        /// The number of hundredths of millimeters (0.01 mm) in an inch. For more information, see GetImagePrefix() method.
        /// </summary>
        private const int HmmPerInch = 2540;

        /// <summary>
        /// The number of twips in an inch. For more information, see GetImagePrefix() method.
        /// </summary>
        private const int TwipsPerInch = 1440;
        #endregion

        #region Fields
        private readonly Graphics _graphics;
        #endregion

        #region Internal class NativeMethods
        internal static class NativeMethods
        {
            #region Enum EmfToWmfBitsFlags
            internal enum EmfToWmfBitsFlags
            {
                // Use the default conversion
                EmfToWmfBitsFlagsDefault = 0x0,
                // Embedded the source of the EMF metafiel within the resulting WMF
                // metafile
                EmfToWmfBitsFlagsEmbedEmf = 0x1,
                // Place a 22-byte header in the resulting WMF file.  The header is
                // required for the metafile to be considered placeable.
                EmfToWmfBitsFlagsIncludePlaceable = 0x2,
                // Don't simulate clipping by using the XOR operator.
                EmfToWmfBitsFlagsNoXorClip = 0x4
            };
            #endregion

            #region Enum ShGetFileInfoConstants
            [Flags]
            internal enum ShGetFileInfoConstants : uint
            {
                /// <summary>
                ///     Get Icon
                /// </summary>
                ShgfiIcon = 0x100,

                /// <summary>
                ///     Get display name
                /// </summary>
                ShgfiDisplayName = 0x200,

                /// <summary>
                ///     Get type name
                /// </summary>
                ShgfiTypeName = 0x400,

                /// <summary>
                ///     Get attributes
                /// </summary>
                ShgfiAttributes = 0x800,

                /// <summary>
                ///     Get icon location
                /// </summary>
                ShgfiIconLocation = 0x1000,

                /// <summary>
                ///     Return exe type
                /// </summary>
                ShgfiExeType = 0x2000,

                /// <summary>
                ///     Get system icon index
                /// </summary>
                ShgfiSysIconIndex = 0x4000,

                /// <summary>
                ///     Put a link overlay on ion
                /// </summary>
                ShgfiLinkOverlay = 0x8000,

                /// <summary>
                ///     Show icon in selected state
                /// </summary>
                ShgfiSelected = 0x10000,

                /// <summary>
                ///     Get only specified attributes
                /// </summary>
                ShgfiAttrSpecified = 0x20000,

                /// <summary>
                ///     Get large icon
                /// </summary>
                ShgfiLargeIcon = 0x0,

                /// <summary>
                ///     Get small icon
                /// </summary>
                ShgfiSmallIcon = 0x1,

                /// <summary>
                ///     Get open icon
                /// </summary>
                ShgfiOpenIcon = 0x2,

                /// <summary>
                ///     Get shell size icon
                /// </summary>
                ShgfiShellIconSize = 0x4,

                /// <summary>
                ///     pszPath is a pidl
                /// </summary>
                ShgfiPidl = 0x8,

                /// <summary>
                ///     Use passed dwFileAttribute
                /// </summary>
                ShgfiUseFileAttributes = 0x10,

                /// <summary>
                ///     Apply the appropriate overlays
                /// </summary>
                ShgfiAddOverlays = 0x000000020,

                /// <summary>
                ///     Get the index of the overlay
                /// </summary>
                ShgfiOverlayIndex = 0x0000000
            }
            #endregion

            #region Struct ShFileInfo
            [StructLayout(LayoutKind.Sequential)]
            internal struct ShFileinfo
            {
                public readonly IntPtr hIcon;
                public readonly IntPtr iIcon;
                public readonly uint dwAttributes;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public readonly string szDisplayName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
                public readonly string szTypeName;
            };
            #endregion

            #region DllImports
            [DllImport("shell32.dll")]
            internal static extern IntPtr SHGetFileInfo(string pszPath,
                uint dwFileAttributes,
                ref ShFileinfo psfi,
                uint cbSizeFileInfo,
                uint uFlags);

            /// <summary>
            ///     Use the EmfToWmfBits function in the GDI+ specification to convert a
            ///     Enhanced Metafile to a Windows Metafile
            /// </summary>
            /// <param name="hEmf">
            ///     A handle to the Enhanced Metafile to be converted
            /// </param>
            /// <param name="bufferSize">
            ///     The size of the buffer used to store the Windows Metafile bits returned
            /// </param>
            /// <param name="buffer">
            ///     An array of bytes used to hold the Windows Metafile bits returned
            /// </param>
            /// <param name="mappingMode">
            ///     The mapping mode of the image.  This control uses MM_ANISOTROPIC.
            /// </param>
            /// <param name="flags">
            ///     Flags used to specify the format of the Windows Metafile returned
            /// </param>
            [DllImport("gdiplus.dll")]
            internal static extern uint GdipEmfToWmfBits(IntPtr hEmf, uint bufferSize, byte[] buffer, int mappingMode, EmfToWmfBitsFlags flags);
            #endregion
        }
        #endregion

        #region Constructor
        internal RtfFileIcon(Graphics graphics)
        {
            _graphics = graphics;
        }
        #endregion

        #region GetImagePrefix
        /// <summary>
        ///     Creates the RTF control string that describes the image being inserted.
        ///     The prefix should have the form ...
        ///     <code>{\pict\wmetafile8\picw[A]\pich[B]\picwgoal[C]\pichgoal[D]</code>
        ///     where ...
        ///     A = current width of the metafile in hundredths of millimeters (0.01mm)
        ///     = Image Width in Inches * Number of (0.01mm) per inch
        ///     = (Image Width in Pixels / Graphics Context's Horizontal Resolution) * 2540
        ///     = (Image Width in Pixels / Graphics.DpiX) * 2540
        ///     B = current height of the metafile in hundredths of millimeters (0.01mm)
        ///     = Image Height in Inches * Number of (0.01mm) per inch
        ///     = (Image Height in Pixels / Graphics Context's Vertical Resolution) * 2540
        ///     = (Image Height in Pixels / Graphics.DpiX) * 2540
        ///     C = target width of the metafile in twips
        ///     = Image Width in Inches * Number of twips per inch
        ///     = (Image Width in Pixels / Graphics Context's Horizontal Resolution) * 1440
        ///     = (Image Width in Pixels / Graphics.DpiX) * 1440
        ///     D = target height of the metafile in twips
        ///     = Image Height in Inches * Number of twips per inch
        ///     = (Image Height in Pixels / Graphics Context's Horizontal Resolution) * 1440
        ///     = (Image Height in Pixels / Graphics.DpiX) * 1440
        /// </summary>
        /// <remarks>
        ///     The Graphics Context's resolution is simply the current resolution at which
        ///     windows is being displayed.  Normally it's 96 dpi.
        ///     According to Ken Howe at pbdr.com, "Twips are screen-independent units
        ///     used to ensure that the placement and proportion of screen elements in
        ///     your screen application are the same on all display systems."
        ///     Units Used
        ///     ----------
        ///     1 Twip = 1/20 Point
        ///     1 Point = 1/72 Inch
        ///     1 Twip = 1/1440 Inch
        ///     1 Inch = 2.54 cm
        ///     1 Inch = 25.4 mm
        ///     1 Inch = 2540 (0.01)mm
        /// </remarks>
        /// <param name="image">
        ///     image which has to be inserted to RTF
        /// </param>
        /// <returns>
        ///     RTF control string that describes the image
        ///     <br/>
        ///         EXAMPLE:
        ///         <code>{\pict\wmetafile8\picw[A]\pich[B]\picwgoal[C]\pichgoal[D]</code>
        /// </returns>
        private string GetImagePrefix(Image image)
        {
            var rtf = string.Empty;

            // Get the horizontal and vertical resolutions at which the object is
            // being displayed
            var graphics = _graphics;

            // The horizontal resolution at which the control is being displayed
            var xDpi = graphics.DpiX;

            // The vertical resolution at which the control is being displayed
            var yDpi = graphics.DpiY;
            
            // Calculate the current width of the image in (0.01)mm
            var picw = (int) Math.Round((image.Width/xDpi)*HmmPerInch);

            // Calculate the current height of the image in (0.01)mm
            var pich = (int) Math.Round((image.Height/yDpi)*HmmPerInch);

            // Calculate the target width of the image in twips
            var picwgoal = (int) Math.Round((image.Width/xDpi)*TwipsPerInch);

            // Calculate the target height of the image in twips
            var pichgoal = (int) Math.Round((image.Height/yDpi)*TwipsPerInch);

            // Append values to RTF string
            rtf = rtf + "{\\pict\\wmetafile8" +
                   "\\picw" + picw +
                   "\\pich" + pich +
                   "\\picwgoal" + picwgoal +
                   "\\pichgoal" + pichgoal + " ";

            return rtf;
        }
        #endregion

        #region GetRtfImage
        /// <summary>
        ///     Wraps the image in an Enhanced Metafile by drawing the image onto the
        ///     graphics context, then converts the Enhanced Metafile to a Windows
        ///     Metafile, and finally appends the bits of the Windows Metafile in HEX
        ///     to a string and returns the string.
        /// </summary>
        /// <param name="image">
        ///     image which has to be converted
        /// </param>
        /// <returns>
        ///     A string containing the bits of a Windows Metafile in HEX
        /// </returns>
        private string GetRtfImage(Image image)
        {
            var rtf = string.Empty;

            // Used to store the enhanced metafile
            MemoryStream stream = null;

            // Used to create the metafile and draw the image
            Graphics graphics = null;

            // The enhanced metafile
            Metafile metaFile = null;

            try
            {
                stream = new MemoryStream();

                // Get a graphics context from the RichTextBox
                using (graphics = _graphics)
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
                    // Draw the image on the Enhanced Metafile
                    graphics.DrawImage(image, new Rectangle(0, 0, image.Width, image.Height));

                // Get the handle of the Enhanced Metafile
                var hEmf = metaFile.GetHenhmetafile();

                // A call to EmfToWmfBits with a null buffer return the size of the
                // buffer need to store the WMF bits.  Use this to get the buffer size.
                var bufferSize = NativeMethods.GdipEmfToWmfBits(hEmf, 0, null, MmAnisotropic,
                    NativeMethods.EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);

                // Create an array to hold the bits
                var buffer = new byte[bufferSize];

                // A call to EmfToWmfBits with a valid buffer copies the bits into the
                // buffer an returns the number of bits in the WMF.  
                NativeMethods.GdipEmfToWmfBits(hEmf, bufferSize, buffer, MmAnisotropic,
                    NativeMethods.EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);

                // Append the bits to the RTF string
                for (var i = 0; i < buffer.Length; ++i)
                    rtf = rtf + String.Format("{0:X2}", buffer[i]);

                return rtf;
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

        #region GetWMeta8Image
        /// <summary>
        ///     returns a string which contains a wmetafile8-Icon string.
        ///     The image is wrapped in a Windows Format Metafile
        /// </summary>
        /// <param name="image">
        ///     image which has to be converted to wmetafile8
        /// </param>
        /// <returns>
        ///     A string which contains a wmetafile8-Icon string.
        /// </returns>
        private string GetWMeta8Image(Image image)
        {
            var rtf = string.Empty;

            // Create the image control string and append it to the RTF string
            rtf = rtf + GetImagePrefix(image);

            // Create the Windows Metafile and append its bytes in HEX format
            rtf = rtf + GetRtfImage(image) + "}";

            return rtf;
        }
        #endregion

        #region GetRTFIcon
        /// <summary>
        ///     This Method is used to return the windows file type icon in rtf form
        /// </summary>
        /// <param name="fileName">
        ///     Out of fileName-extenion the file type icon is extracted
        ///     <br/>
        ///         EXAMPLE:
        ///         fileName = "Hello.doc" ---> Result will be the icon for MS Word document
        /// </param>
        /// <returns>
        ///     A String containing a icon in WMeta8 Icon format (RTF Control string).
        /// </returns>
        public string GetRTFIcon(string fileName)
        {
            var shinfo = new NativeMethods.ShFileinfo();
            NativeMethods.SHGetFileInfo(fileName, 0, ref shinfo,
                (uint) Marshal.SizeOf(shinfo),
                (uint)(NativeMethods.ShGetFileInfoConstants.ShgfiIcon |
                        NativeMethods.ShGetFileInfoConstants.ShgfiLargeIcon |
                        NativeMethods.ShGetFileInfoConstants.ShgfiUseFileAttributes));

            var fTypeIcon = Icon.FromHandle(shinfo.hIcon);
            Image imgIcon = fTypeIcon.ToBitmap();
            var rtfString = GetWMeta8Image(imgIcon);

            return rtfString;
        }
        #endregion
    }
}