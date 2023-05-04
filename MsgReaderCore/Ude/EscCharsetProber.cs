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

namespace MsgReader.Ude;

internal class EscCharsetProber : CharsetProber
{
    #region Fields
    private const int CharsetsNum = 4;
    private string _detectedCharset;
    private readonly CodingStateMachine[] _codingSm;
    private int _activeSm;
    #endregion

    #region Constructor
    internal EscCharsetProber()
    {
        _codingSm = new CodingStateMachine[CharsetsNum];
        _codingSm[0] = new CodingStateMachine(new HzsmModel());
        _codingSm[1] = new CodingStateMachine(new Iso2022CnsmModel());
        _codingSm[2] = new CodingStateMachine(new Iso2022JpsmModel());
        _codingSm[3] = new CodingStateMachine(new Iso2022KrsmModel());
        Reset();
    }
    #endregion

    #region Reset
    public sealed override void Reset()
    {
        State = ProbingState.Detecting;
        for (var i = 0; i < CharsetsNum; i++)
            _codingSm[i].Reset();
        _activeSm = CharsetsNum;
        _detectedCharset = null;
    }
    #endregion

    #region HandleData
    public override ProbingState HandleData(byte[] buf, int offset, int len)
    {
        var max = offset + len;

        for (var i = offset; i < max && State == ProbingState.Detecting; i++)
        for (var j = _activeSm - 1; j >= 0; j--)
        {
            // byte is feed to all active state machine
            var codingState = _codingSm[j].NextState(buf[i]);
            if (codingState == SmModel.Error)
            {
                // got negative answer for this state machine, make it inactive
                _activeSm--;
                if (_activeSm == 0)
                {
                    State = ProbingState.NotMe;
                    return State;
                }

                if (j != _activeSm)
                    (_codingSm[_activeSm], _codingSm[j]) = (_codingSm[j], _codingSm[_activeSm]);
            }
            else if (codingState == SmModel.ItsMe)
            {
                State = ProbingState.FoundIt;
                _detectedCharset = _codingSm[j].ModelName;
                return State;
            }
        }

        return State;
    }
    #endregion

    #region GetCharsetName
    public override string GetCharsetName()
    {
        return _detectedCharset;
    }
    #endregion

    #region GetConfidence
    public override float GetConfidence()
    {
        return 0.99f;
    }
    #endregion
}