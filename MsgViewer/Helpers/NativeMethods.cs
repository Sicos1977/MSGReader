using System;
using System.Runtime.InteropServices;
// ReSharper disable InconsistentNaming

namespace MsgViewer.Helpers
{
    #region Public static class NativeMethods
    public static class NativeMethods
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
}
