//
// StreamHelpers.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2022 Magic-Sessions. (www.magic-sessions.com)
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
using Microsoft.IO;

namespace MsgReader.Helpers
{
    /// <summary>
    /// A class with stream helper methods
    /// </summary>
    public static class StreamHelpers
    {
        #region Consts
        private const int BlockSize = 1024;
        private const int LargeBufferMultiple = 1024 * 1024;
        private const int MaxBufferSize = 16 * LargeBufferMultiple;
        #endregion

        #region Fields
        private static RecyclableMemoryStreamManager _manager;
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the <see cref="RecyclableMemoryStreamManager"/>
        /// </summary>
        /// <remarks>
        /// https://github.com/microsoft/Microsoft.IO.RecyclableMemoryStream
        /// </remarks>
        public static RecyclableMemoryStreamManager Manager
        {
            get
            {
                if (_manager != null)
                    return _manager;

                _manager = new RecyclableMemoryStreamManager(BlockSize, LargeBufferMultiple, MaxBufferSize)
                {
                    //_manager.GenerateCallStacks = true;
                    AggressiveBufferReturn = true,
                    MaximumFreeLargePoolBytes = MaxBufferSize * 4,
                    MaximumFreeSmallPoolBytes = 100 * BlockSize
                };

                return _manager;
            }

            set => _manager = value;
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