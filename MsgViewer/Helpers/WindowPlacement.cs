using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace MsgViewer.Helpers
{
    /// <summary>
    /// This class is a P/INVOKE wrapper to the Window placement function.
    /// http://msdn.microsoft.com/en-us/library/windows/desktop/ms632611%28v=vs.85%29.aspx
    /// </summary>
    public static class WindowPlacement
    {
        // ReSharper disable InconsistentNaming
        #region Fields
        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        private static readonly Encoding Encoding = new UTF8Encoding();
        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof (NativeMethods.WINDOWPLACEMENT));
        #endregion

        #region Public static class NativeMethods
        public class NativeMethods
        {
            #region Structures
            /// <summary>
            /// RECT structure required by WINDOWPLACEMENT structure
            /// </summary>
            [Serializable]
            [StructLayout(LayoutKind.Sequential)]
            public struct RECT
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;

                public RECT(int left, int top, int right, int bottom)
                {
                    Left = left;
                    Top = top;
                    Right = right;
                    Bottom = bottom;
                }
            }

            /// <summary>
            /// POINT structure required by WINDOWPLACEMENT structure
            /// </summary>
            [Serializable]
            [StructLayout(LayoutKind.Sequential)]
            public struct POINT
            {
                public int X;
                public int Y;

                public POINT(int x, int y)
                {
                    X = x;
                    Y = y;
                }
            }

            /// <summary>
            /// WINDOWPLACEMENT stores the position, size, and state of a window
            /// </summary>
            [Serializable]
            [StructLayout(LayoutKind.Sequential)]
            public struct WINDOWPLACEMENT
            {
                public int length;
                public int flags;
                public int showCmd;
                public POINT minPosition;
                public POINT maxPosition;
                public RECT normalPosition;
            }
            #endregion

            [DllImport("user32.dll")]
            public static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WINDOWPLACEMENT lpwndpl);

            [DllImport("user32.dll")]
            public static extern bool GetWindowPlacement(IntPtr hWnd, out WINDOWPLACEMENT lpwndpl);
        }
        #endregion

        #region SetPlacement
        /// <summary>
        /// Sets the position of the window
        /// </summary>
        /// <param name="windowHandle"></param>
        /// <param name="placementXml"></param>
        public static void SetPlacement(IntPtr windowHandle, string placementXml)
        {
            if (string.IsNullOrEmpty(placementXml))
                return;

            var xmlBytes = Encoding.GetBytes(placementXml);

            try
            {
                NativeMethods.WINDOWPLACEMENT placement;
                using (var memoryStream = new MemoryStream(xmlBytes))
                    placement = (NativeMethods.WINDOWPLACEMENT) Serializer.Deserialize(memoryStream);

                placement.length = Marshal.SizeOf(typeof (NativeMethods.WINDOWPLACEMENT));
                placement.flags = 0;
                placement.showCmd = (placement.showCmd == SW_SHOWMINIMIZED ? SW_SHOWNORMAL : placement.showCmd);
                NativeMethods.SetWindowPlacement(windowHandle, ref placement);
            }
            catch (InvalidOperationException)
            {
                // Parsing placement XML failed. Fail silently.
            }
        }
        #endregion

        #region GetPlacement
        /// <summary>
        /// Returns the position of the window
        /// </summary>
        /// <param name="windowHandle"></param>
        /// <returns></returns>
        public static string GetPlacement(IntPtr windowHandle)
        {
            NativeMethods.WINDOWPLACEMENT placement;
            NativeMethods.GetWindowPlacement(windowHandle, out placement);

            using (var memoryStream = new MemoryStream())
            {
                using (var xmlTextWriter = new XmlTextWriter(memoryStream, Encoding.UTF8))
                {
                    Serializer.Serialize(xmlTextWriter, placement);
                    var xmlBytes = memoryStream.ToArray();
                    return Encoding.GetString(xmlBytes);
                }
            }
        }
        #endregion
        // ReSharper restore InconsistentNaming
    }
}