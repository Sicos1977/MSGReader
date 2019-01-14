//
// StreamHelpers.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System.IO;

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