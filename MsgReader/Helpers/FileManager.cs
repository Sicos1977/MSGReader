using System;
using System.Globalization;
using System.IO;
using System.Linq;

/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader.Helpers
{
    /// <summary>
    /// This class contains file management functions that are not available in the .NET framework
    /// </summary>
    internal static class FileManager
    {
        #region Consts
        /// <summary>
        /// The max path length in Windows
        /// </summary>
        private const int MaxPath = 248;
        #endregion

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

        #region ValidateLongFileName
        /// <summary>
        /// Validates the length of <paramref name="fileName"/>, when this is longer then <see cref="MaxPath"/> chars it will be truncated.
        /// </summary>
        /// <param name="fileName">The filename with path</param>
        /// <param name="extraTruncateSize">Optional extra truncate size, when not used the filename is truncated until it fits</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">Raised when no path or file name is given in the <paramref name="fileName"/></exception>
        /// <exception cref="PathTooLongException">Raised when it is not possible to truncate the <paramref name="fileName"/></exception>
        public static string ValidateLongFileName(string fileName, int extraTruncateSize = -1)
        {
            var fileNameWithoutExtension = GetFileNameWithoutExtension(fileName);

            if (string.IsNullOrWhiteSpace(fileNameWithoutExtension))
                throw new ArgumentException(@"No file name is given, e.g. c:\temp\temp.txt", "fileName");

            var extension = GetExtension(fileName);

            if (string.IsNullOrWhiteSpace(extension))
                extension = string.Empty;

            var path = GetDirectoryName(fileName);

            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException(@"No path is given, e.g. c:\temp\temp.txt", "fileName");

            path = CheckForBackSlash(path);

            if (fileName.Length <= MaxPath)
                return fileName;

            var maxFileNameLength = MaxPath - path.Length - extension.Length;
            if (extraTruncateSize != -1)
                maxFileNameLength -= extraTruncateSize;

            if (maxFileNameLength < 1)
                throw new PathTooLongException("Unable the truncate the fileName '" + fileName + "', current size '" +
                                               fileName.Length + "'");

            return path + fileNameWithoutExtension.Substring(0, maxFileNameLength) + extension;
        }
        #endregion

        #region GetExtension
        /// <summary>
        /// Returns the extension of the specified <paramref name="path"/> string
        /// </summary>
        /// <param name="path">The path of the file</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">Raised when no path is given</exception>
        public static string GetExtension(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("path");

            var splittedPath = path.Split(Path.DirectorySeparatorChar);
            var fileName = splittedPath[splittedPath.Length - 1];

            var index = fileName.LastIndexOf(".", StringComparison.Ordinal);

            return index == -1
                ? string.Empty
                : fileName.Substring(fileName.LastIndexOf(".", StringComparison.Ordinal), fileName.Length - index);
        }
        #endregion

        #region GetFileNameWithoutExtension
        /// <summary>
        /// Returns the file name of the specified <paramref name="path"/> string without the extension
        /// </summary>
        /// <param name="path">The path of the file</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static string GetFileNameWithoutExtension(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException(@"No path given", "path");

            var splittedPath = path.Split(Path.DirectorySeparatorChar);
            var fileName = splittedPath[splittedPath.Length - 1];
            return !fileName.Contains(".")
                ? fileName
                : fileName.Substring(0, fileName.LastIndexOf(".", StringComparison.Ordinal));
        }
        #endregion

        #region GetDirectoryName
        /// <summary>
        /// Returns the directory information for the specified <paramref name="path"/> string
        /// </summary>
        /// <param name="path">The path of a file or directory</param>
        /// <returns></returns>
        public static string GetDirectoryName(string path)
        {
            //GetDirectoryName('C:\MyDir\MySubDir\myfile.ext') returns 'C:\MyDir\MySubDir'
            //GetDirectoryName('C:\MyDir\MySubDir') returns 'C:\MyDir'
            //GetDirectoryName('C:\MyDir\') returns 'C:\MyDir'
            //GetDirectoryName('C:\MyDir') returns 'C:\'
            //GetDirectoryName('C:\') returns ''

            var splittedPath = path.Split(Path.DirectorySeparatorChar);

            if (splittedPath.Length <= 1)
                return string.Empty;

            var result = splittedPath[0];

            for (var i = 1; i < splittedPath.Length - 1; i++)
                result += Path.DirectorySeparatorChar + splittedPath[i];

            return result;
        }
        #endregion

        #region FileExistsMakeNew
        /// <summary>
        /// Checks if a file already exists and if so adds a number until the file is unique
        /// </summary>
        /// <param name="fileName">The file to check</param>
        /// <param name="validateLongFileName">When true validation will be performed on the max path lengt</param>
        /// <param name="extraTruncateSize"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">Raised when no path or file name is given in the <paramref name="fileName"/></exception>
        /// <exception cref="PathTooLongException">Raised when it is not possible to truncate the <paramref name="fileName"/></exception>
        public static string FileExistsMakeNew(string fileName, bool validateLongFileName = true, int extraTruncateSize = -1)
        {
            var extension = GetExtension(fileName);
            var path = CheckForBackSlash(GetDirectoryName(fileName));

            var tempFileName = validateLongFileName ? ValidateLongFileName(fileName, extraTruncateSize) : fileName;

            var i = 2;
            while (File.Exists(tempFileName))
            {
                tempFileName = validateLongFileName ? ValidateLongFileName(tempFileName, extraTruncateSize) : tempFileName;
                var fileNameWithoutExtension = GetFileNameWithoutExtension(tempFileName);
                tempFileName = path + fileNameWithoutExtension + "_" + i + extension;
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