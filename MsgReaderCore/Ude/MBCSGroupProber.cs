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
 *          Kohei TAKETA <k-tak@void.in> (Java port)
 *          Rudi Pettazzi <rudi.pettazzi@gmail.com> (C# port)
 *
 * Refactoring to the code done by Kees van Spelde so that it works in this project
 * Copyright (c) 2023 Magic-Sessions. (www.magic-sessions.com)
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

using System.Text;

namespace MsgReader.Ude;

/// <summary>
///     Multi-byte charsets probers
/// </summary>
internal class MbcsGroupProber : CharsetProber
{
    #region Consts
    private const int ProbersNum = 7;
    #endregion

    #region Fields
    private static readonly string[] ProberName = { "UTF8", "SJIS", "EUCJP", "GB18030", "EUCKR", "Big5", "EUCTW" };
    private readonly CharsetProber[] _probers = new CharsetProber[ProbersNum];
    private readonly bool[] _isActive = new bool[ProbersNum];
    private int _bestGuess;
    private int _activeNum;
    #endregion

    #region Constructor
    internal MbcsGroupProber()
    {
        _probers[0] = new UTF8Prober();
        _probers[1] = new SJISProber();
        _probers[2] = new EucjpProber();
        _probers[3] = new Gb18030Prober();
        _probers[4] = new EuckrProber();
        _probers[5] = new Big5Prober();
        _probers[6] = new EuctwProber();
        Reset();
    }
    #endregion

    #region GetCharsetName
    public override string GetCharsetName()
    {
        if (_bestGuess != -1) return _probers[_bestGuess].GetCharsetName();
        GetConfidence();
        if (_bestGuess == -1)
            _bestGuess = 0;

        return _probers[_bestGuess].GetCharsetName();
    }
    #endregion

    #region Reset
    public sealed override void Reset()
    {
        _activeNum = 0;
        for (var i = 0; i < _probers.Length; i++)
            if (_probers[i] != null)
            {
                _probers[i].Reset();
                _isActive[i] = true;
                ++_activeNum;
            }
            else
            {
                _isActive[i] = false;
            }

        _bestGuess = -1;
        State = ProbingState.Detecting;
    }
    #endregion

    #region HandleData
    public override ProbingState HandleData(byte[] buf, int offset, int len)
    {
        // do filtering to reduce load to probers
        var highbyteBuf = new byte[len];
        var hptr = 0;
        //assume previous is not ascii, it will do no harm except add some noise
        var keepNext = true;
        var max = offset + len;

        for (var i = offset; i < max; i++)
            if ((buf[i] & 0x80) != 0)
            {
                highbyteBuf[hptr++] = buf[i];
                keepNext = true;
            }
            else
            {
                //if previous is highbyte, keep this even it is a ASCII
                if (keepNext)
                {
                    highbyteBuf[hptr++] = buf[i];
                    keepNext = false;
                }
            }

        for (var i = 0; i < _probers.Length; i++)
        {
            if (!_isActive[i])
                continue;
            var st = _probers[i].HandleData(highbyteBuf, 0, hptr);
            if (st == ProbingState.FoundIt)
            {
                _bestGuess = i;
                State = ProbingState.FoundIt;
                break;
            }

            if (st != ProbingState.NotMe) continue;
            _isActive[i] = false;
            _activeNum--;
            if (_activeNum > 0) continue;
            State = ProbingState.NotMe;
            break;
        }

        return State;
    }
    #endregion

    #region GetConfidence
    public override float GetConfidence()
    {
        var bestConf = 0.0f;

        switch (State)
        {
            case ProbingState.FoundIt:
                return 0.99f;
            case ProbingState.NotMe:
                return 0.01f;
        }

        for (var i = 0; i < ProbersNum; i++)
        {
            if (!_isActive[i])
                continue;
            var cf = _probers[i].GetConfidence();
            if (!(bestConf < cf)) continue;
            bestConf = cf;
            _bestGuess = i;
        }

        return bestConf;
    }
    #endregion

    #region DumpStatus
    public override string DumpStatus()
    {
        GetConfidence();

        var sb = new StringBuilder();
        for (var i = 0; i < ProbersNum; i++)
        {
            if (sb.Length != 0) sb.AppendLine();

            if (!_isActive[i])
            {
                sb.Append($"MBCS inactive: {ProberName[i]} (confidence is too low).");
            }
            else
            {
                var cf = _probers[i].GetConfidence();
                sb.Append($"MBCS {cf}: [{ProberName[i]}]");
            }
        }

        return sb.ToString();
    }
    #endregion
}