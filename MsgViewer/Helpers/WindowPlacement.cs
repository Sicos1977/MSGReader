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
        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof(NativeMethods.WINDOWPLACEMENT));
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

            try
            {
                var xmlBytes = Encoding.GetBytes(placementXml);
                NativeMethods.WINDOWPLACEMENT placement;
                using (var memoryStream = new MemoryStream(xmlBytes))
                    placement = (NativeMethods.WINDOWPLACEMENT)Serializer.Deserialize(memoryStream);

                placement.length = Marshal.SizeOf(typeof(NativeMethods.WINDOWPLACEMENT));
                placement.flags = 0;
                placement.showCmd = (placement.showCmd == SW_SHOWMINIMIZED ? SW_SHOWNORMAL : placement.showCmd);
                NativeMethods.SetWindowPlacement(windowHandle, ref placement);
            }
            catch
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
            NativeMethods.GetWindowPlacement(windowHandle, out var placement);

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