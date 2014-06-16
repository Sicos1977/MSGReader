using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace MsgViewer
{
    public static class WindowPlacement
    {
        #region Fields
        // ReSharper disable InconsistentNaming
        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        // ReSharper restore InconsistentNaming
        private static readonly Encoding Encoding = new UTF8Encoding();
        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof (NativeMethods.Windowplacement));
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
            public struct Rect
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;

                public Rect(int left, int top, int right, int bottom)
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
            public struct Point
            {
                public int X;
                public int Y;

                public Point(int x, int y)
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
            public struct Windowplacement
            {
                public int length;
                public int flags;
                public int showCmd;
                public Point minPosition;
                public Point maxPosition;
                public Rect normalPosition;
            }
            #endregion

            [DllImport("user32.dll")]
            public static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref Windowplacement lpwndpl);

            [DllImport("user32.dll")]
            public static extern bool GetWindowPlacement(IntPtr hWnd, out Windowplacement lpwndpl);
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
                NativeMethods.Windowplacement placement;
                using (var memoryStream = new MemoryStream(xmlBytes))
                    placement = (NativeMethods.Windowplacement) Serializer.Deserialize(memoryStream);

                placement.length = Marshal.SizeOf(typeof (NativeMethods.Windowplacement));
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
            NativeMethods.Windowplacement placement;
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
    }
}