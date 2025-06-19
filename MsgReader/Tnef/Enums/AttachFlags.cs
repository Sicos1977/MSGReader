﻿//
// TnefAttachFlags.cs
//
// Author: Jeffrey Stedfast <jestedfa@microsoft.com>
//
// Copyright (c) 2013-2022 .NET Foundation and Contributors
//
// Refactoring to the code done by Kees van Spelde so that it works in this project
// Copyright (c) 2023 Magic-Sessions. (www.magic-sessions.com)
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

using System;

namespace MsgReader.Tnef.Enums;

/// <summary>
///     The TNEF attach flags.
/// </summary>
/// <remarks>
///     The <see cref="AttachFlags" /> enum contains a list of possible values for
///     the <see cref="PropertyId.AttachFlags" /> property.
/// </remarks>
[Flags]
internal enum AttachFlags
{
    /// <summary>
    ///     No AttachFlags set.
    /// </summary>
    None = 0,

    /// <summary>
    ///     The attachment is invisible in HTML bodies.
    /// </summary>
    InvisibleInHtml = 1,

    /// <summary>
    ///     The attachment is invisible in RTF bodies.
    /// </summary>
    InvisibleInRtf = 2,

    /// <summary>
    ///     The attachment is referenced (and rendered) by the HTML body.
    /// </summary>
    RenderedInBody = 4
}