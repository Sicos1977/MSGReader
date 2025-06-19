﻿//
// LayerInfo.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2024 Kees van Spelde. (www.magic-sessions.com)
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

namespace MsgReader.Rtf;

/// <summary>
///     Rtf raw lay
/// </summary>
internal class LayerInfo
{
    #region Fields
    private int _ucValue;
    #endregion

    #region Properties
    public int UcValue
    {
        get => _ucValue;
        set
        {
            _ucValue = value;
            UcValueCount = 0;
        }
    }

    public int UcValueCount { get; set; }
    #endregion

    #region CheckUcValueCount
    /// <summary>
    ///     Checks if Uc value count is greater then zero
    /// </summary>
    /// <returns>True when greater</returns>
    public bool CheckUcValueCount()
    {
        UcValueCount--;
        return UcValueCount < 0;
    }
    #endregion

    #region Clone
    public LayerInfo Clone()
    {
        return (LayerInfo)MemberwiseClone();
    }
    #endregion
}