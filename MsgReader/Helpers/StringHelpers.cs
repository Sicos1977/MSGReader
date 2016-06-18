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
        /// <returns></returns>
        public static string ReadNullTerminatedString(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            var chr = binaryReader.ReadChar();
            while (chr != 0)
            {
                stringBuilder.Append(chr);
                chr = binaryReader.ReadChar();
            }

            return stringBuilder.ToString();
        }
        #endregion
    }
}
