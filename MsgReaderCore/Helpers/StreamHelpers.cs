using System.IO;

/*
   Copyright 2015 - 2016 Kees van Spelde

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
    internal static class StreamHelpers
    {
        #region ToByteArray
        /// <summary>
        ///     Returns the stream as an byte array
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        internal static byte[] ToByteArray(this Stream input)
        {
            using (var memoryStream = new MemoryStream())
            {
                input.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }
        #endregion

        #region Eos
        /// <summary>
        ///     Returns true when the end of the <see cref="BinaryReader.BaseStream" /> has been reached
        /// </summary>
        /// <param name="binaryReader"></param>
        /// <returns></returns>
        internal static bool Eos(this BinaryReader binaryReader)
        {
            try
            {
                return binaryReader.BaseStream.Position >= binaryReader.BaseStream.Length;
            }
            catch (IOException)
            {
                return true;
            }
        }
        #endregion
    }
}