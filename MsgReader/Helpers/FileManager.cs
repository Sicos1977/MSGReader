using System;
using System.Globalization;
using System.IO;
using System.Linq;

namespace DocumentServices.Modules.Readers.MsgReader.Helpers
{
    /// <summary>
    /// This class contains file management functions that are not available in the .NET framework
    /// </summary>
    internal static class FileManager
    {
        #region CheckForBackSlash
        /// <summary>
        /// Check if there is a backslash at the end of the string and if not add it
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public static string CheckForBackSlash(string line)
        {
            if (line.EndsWith("\\"))
                return line;

            return line + "\\";
        }
        #endregion

        #region FileExistsMakeNew
        /// <summary>
        /// Checks if a file already exists and if so adds a number until the file is unique
        /// </summary>
        /// <param name="fileName">The file to check</param>
        /// <returns></returns>
        public static string FileExistsMakeNew(string fileName)
        {
            var i = 2;
            var path = CheckForBackSlash(Path.GetDirectoryName(fileName));

            var tempFileName = fileName;

            while (File.Exists(tempFileName))
            {
                tempFileName = path + Path.GetFileNameWithoutExtension(fileName) + "_" + i + Path.GetExtension(fileName);
                i += 1;
            }

            return tempFileName;
        }
        #endregion

        #region RemoveInvalidFileNameChars
        /// <summary>
        /// Removes illegal filename characters
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string RemoveInvalidFileNameChars(string fileName)
        {
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c.ToString(CultureInfo.InvariantCulture), string.Empty));
        }
        #endregion

        #region GetFileSizeString
        /// <summary>
        /// Gives the size of a file in Windows format (GB, MB, KB, Bytes)
        /// </summary>
        /// <param name="bytes">Filesize in bytes</param>
        /// <returns></returns>
        public static string GetFileSizeString(double bytes)
        {
            var size = "0 Bytes";
            if (bytes >= 1073741824.0)
                size = String.Format(CultureInfo.InvariantCulture, "{0:##.##}", bytes / 1073741824.0) + " GB";
            else if (bytes >= 1048576.0)
                size = String.Format(CultureInfo.InvariantCulture, "{0:##.##}", bytes / 1048576.0) + " MB";
            else if (bytes >= 1024.0)
                size = String.Format(CultureInfo.InvariantCulture, "{0:##.##}", bytes / 1024.0) + " KB";
            else if (bytes > 0 && bytes < 1024.0)
                size = bytes + " Bytes";

            return size;
        }
        #endregion
    }
}