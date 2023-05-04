/* ***** BEGIN LICENSE BLOCK *****
 * Version: MPL 1.1/GPL 2.0/LGPL 2.1
 *
 * The contents of this file are subject to the Mozilla Public License Version
 * 1.1 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 * http://www.mozilla.org/MPL/
 *
 * Software distributed under the License is distributed on an "AS IS" basis,
 * WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
 * for the specific language governing rights and limitations under the
 * License.
 *
 * The Original Code is Mozilla Universal charset detector code.
 *
 * The Initial Developer of the Original Code is
 * Netscape Communications Corporation.
 * Portions created by the Initial Developer are Copyright (C) 2001
 * the Initial Developer. All Rights Reserved.
 *
 * Contributor(s):
 *          Shy Shalom <shooshX@gmail.com>
 *          Rudi Pettazzi <rudi.pettazzi@gmail.com> (C# port)
 * 
 * Alternatively, the contents of this file may be used under the terms of
 * either the GNU General Public License Version 2 or later (the "GPL"), or
 * the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
 * in which case the provisions of the GPL or the LGPL are applicable instead
 * of those above. If you wish to allow use of your version of this file only
 * under the terms of either the GPL or the LGPL, and not to allow others to
 * use your version of this file under the terms of the MPL, indicate your
 * decision by deleting the provisions above and replace them with the notice
 * and other provisions required by the GPL or the LGPL. If you do not delete
 * the provisions above, a recipient may use your version of this file under
 * the terms of any one of the MPL, the GPL or the LGPL.
 *
 * ***** END LICENSE BLOCK ***** */

using System.IO;

namespace MsgReader.Ude;

public enum ProbingState
{
    Detecting = 0, // no sure answer yet, but caller can ask for confidence
    FoundIt = 1, // positive answer
    NotMe = 2 // negative answer 
}

internal abstract class CharsetProber
{
    #region Consts
    protected const float ShortcutThreshold = 0.95F;

    // ASCII codes
    private const byte Space = 0x20;
    private const byte CapitalA = 0x41;
    private const byte CapitalZ = 0x5A;
    private const byte SmallA = 0x61;
    private const byte SmallZ = 0x7A;
    private const byte LessThan = 0x3C;
    private const byte GreaterThan = 0x3E;
    #endregion

    #region Fields
    protected ProbingState State;
    #endregion

    #region HandleData
    /// <summary>
    ///     Feed data to the prober
    /// </summary>
    /// <param name="buf">a buffer</param>
    /// <param name="offset">offset into buffer</param>
    /// <param name="len">number of bytes available into buffer</param>
    /// <returns>
    ///     A <see cref="ProbingState" />
    /// </returns>
    public abstract ProbingState HandleData(byte[] buf, int offset, int len);
    #endregion

    #region Abstract methods
    /// <summary>
    ///     Reset prober state
    /// </summary>
    public abstract void Reset();

    public abstract string GetCharsetName();

    public abstract float GetConfidence();
    #endregion

    #region GetState
    public virtual ProbingState GetState()
    {
        return State;
    }
    #endregion

    #region SetOption
    public virtual void SetOption()
    {
    }
    #endregion

    #region DumpStatus
    public virtual string DumpStatus()
    {
        return string.Empty;
    }
    #endregion

    #region FilterWithoutEnglishLetters
    /// <summary>
    ///     Helper functions used in the Latin1 and Group probers
    /// </summary>
    /// <returns>filtered buffer</returns>
    protected static byte[] FilterWithoutEnglishLetters(byte[] buf, int offset, int len)
    {
        using var ms = new MemoryStream(buf.Length);
        var meetMsb = false;
        var max = offset + len;
        var prev = offset;
        var cur = offset;

        while (cur < max)
        {
            var b = buf[cur];

            if ((b & 0x80) != 0)
            {
                meetMsb = true;
            }
            else if (b < CapitalA || (b > CapitalZ && b < SmallA)
                                  || b > SmallZ)
            {
                if (meetMsb && cur > prev)
                {
                    ms.Write(buf, prev, cur - prev);
                    ms.WriteByte(Space);
                    meetMsb = false;
                }

                prev = cur + 1;
            }

            cur++;
        }

        if (meetMsb && cur > prev)
            ms.Write(buf, prev, cur - prev);
        ms.SetLength(ms.Position);
        var result = ms.ToArray();

        return result;
    }
    #endregion

    #region FilterWithEnglishLetters
    /// <summary>
    ///     Do filtering to reduce load to probers (Remove ASCII symbols,
    ///     collapse spaces). This filter applies to all scripts which contain
    ///     both English characters and upper ASCII characters.
    /// </summary>
    /// <returns>a filtered copy of the input buffer</returns>
    protected static byte[] FilterWithEnglishLetters(byte[] buf, int offset, int len)
    {
        using var ms = new MemoryStream(buf.Length);
        var inTag = false;
        var max = offset + len;
        var prev = offset;
        var cur = offset;

        while (cur < max)
        {
            var b = buf[cur];

            inTag = b switch
            {
                GreaterThan => false,
                LessThan => true,
                _ => inTag
            };

            // it's ascii, but it's not a letter
            if ((b & 0x80) == 0 && (b < CapitalA || b > SmallZ
                                                 || (b > CapitalZ && b < SmallA)))
            {
                if (cur > prev && !inTag)
                {
                    ms.Write(buf, prev, cur - prev);
                    ms.WriteByte(Space);
                }

                prev = cur + 1;
            }

            cur++;
        }

        // If the current segment contains more than just a symbol 
        // and it is not inside a tag then keep it.
        if (!inTag && cur > prev)
            ms.Write(buf, prev, cur - prev);
        ms.SetLength(ms.Position);
        var result = ms.ToArray();

        return result;
    }
    #endregion
}