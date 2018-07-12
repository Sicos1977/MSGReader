using System.IO;
using System.Text;

namespace MsgReader.Helpers
{
    internal static class Strings
    {
        #region ReadNullTerminatedString
        /// <summary>
        ///     Reads from the <paramref name="binaryReader"/> until a null terminated char is read
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <param name="unicode"></param>
        /// <returns></returns>
        public static string ReadNullTerminatedString(BinaryReader binaryReader, bool unicode)
        {
            return unicode ? ReadNullTerminatedUnicodeString(binaryReader) : ReadNullTerminatedAsciiString(binaryReader);
        }
        #endregion

        #region ReadNullTerminatedString
        /// <summary>
        ///     Reads from the <paramref name="binaryReader"/> until a null terminated char is read
        /// </summary>
        /// <param name="binaryReader">The <see cref="BinaryReader" /></param>
        /// <returns></returns>
        public static string ReadNullTerminatedAsciiString(BinaryReader binaryReader)
        {
            var result = new MemoryStream();

            var b = binaryReader.ReadByte();
            while (b != 0)
            {
                result.WriteByte(b);
                b = binaryReader.ReadByte();
            }

            return Encoding.ASCII.GetString(result.ToArray());
        }
        #endregion

        #region ReadNullTerminatedUnicodeString
        /// <summary>
        ///     Reads from the <paramref name="binaryReader"/> until a null terminated char is read
        /// </summary>
        /// <param name="binaryReader">The <see cref="BinaryReader" /></param>
        /// <returns></returns>
        public static string ReadNullTerminatedUnicodeString(BinaryReader binaryReader)
        {
            var result = new MemoryStream();

            var b = binaryReader.ReadBytes(2);
            while (b[0] != 0 && b[1] != 0)
            {
                result.WriteByte(b[0]);
                result.WriteByte(b[2]);
                b = binaryReader.ReadBytes(2);
            }

            return Encoding.Unicode.GetString(result.ToArray());
        }
        #endregion
    }
}
