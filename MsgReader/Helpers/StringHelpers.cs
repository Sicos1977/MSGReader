using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MsgReader.Helpers
{
    internal static class StringHelpers
    {
        #region ReadNullTerminatedString
        /// <summary>
        ///     Reads from the <see cref="binaryReader" /> until a null terminated char is read
        /// </summary>
        /// <param name="binaryReader">The <see cref="BinaryReader" /></param>
        /// <param name="unicode">When set to <c>true</c> then the string has to be read as unicode</param>
        /// <returns></returns>
        public static string ReadNullTerminatedString(BinaryReader binaryReader, bool unicode)
        {
            var result = new MemoryStream();
            var b = binaryReader.ReadByte();
            while (b != 0)
            {
                result.WriteByte(b);
                b = binaryReader.ReadByte();
            }

            return unicode
                ? Encoding.Unicode.GetString(result.ToArray())
                : Encoding.ASCII.GetString(result.ToArray());
        }
        #endregion
    }
}
