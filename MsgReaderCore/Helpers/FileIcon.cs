//
// FileIcon.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace MsgReader.Helpers
{
    /// <summary>
    /// This class can be used to return a picture of an icon that is coupled to a specific file type
    /// </summary>
    internal class FileIcon
    {
        #region Internal class NativeMethods
        internal class NativeMethods
        {
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
                public IntPtr hIcon;
                public IntPtr iIcon;
                public uint dwAttributes;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public readonly string szDisplayName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
                public readonly string szTypeName;
            };
            #endregion

            #region DllImports
            [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA2101:SpecifyMarshalingForPInvokeStringArguments", MessageId = "0"), DllImport("shell32.dll")]
            internal static extern IntPtr SHGetFileInfo(
                string pszPath,
                uint dwFileAttributes,
                ref ShFileinfo psfi,
                uint cbSizeFileInfo,
                uint uFlags);
            #endregion
        }
        #endregion

        // TODO: Check if this works in .NET STANDARD 2

        #region GetFileIcon
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
        public static Image GetFileIcon(string fileName)
        {
            var shinfo = new NativeMethods.ShFileinfo();
            NativeMethods.SHGetFileInfo(fileName, 0, ref shinfo,
                (uint) Marshal.SizeOf(shinfo),
                (uint) (NativeMethods.ShGetFileInfoConstants.ShgfiIcon |
                        NativeMethods.ShGetFileInfoConstants.ShgfiLargeIcon |
                        NativeMethods.ShGetFileInfoConstants.ShgfiUseFileAttributes));

            
            var fTypeIcon = Icon.FromHandle(shinfo.hIcon);
            return fTypeIcon.ToBitmap();
        }
        #endregion
    }
}