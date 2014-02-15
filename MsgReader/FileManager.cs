using System.IO;
using System.Linq;

namespace DocumentServices.Modules.Readers.MsgReader
{
    /// <summary>
    /// File management functions that are not available in the .NET framework
    /// </summary>
    internal static class FileManager
    {
        #region CheckForSlash
        /// <summary>
        /// Check if there is a slash at the end of the string and if not add it
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public static string CheckForSlash(string line)
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
            var path = CheckForSlash(Path.GetDirectoryName(fileName));

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
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c.ToString(), string.Empty));
        }
        #endregion
    }
}