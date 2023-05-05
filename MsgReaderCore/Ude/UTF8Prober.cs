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

internal class Utf8Prober : CharsetProber
{
    #region Fields
    private const float OneCharProb = 0.50f;
    private readonly CodingStateMachine _codingSm;
    private int _numOfMbChar;
    #endregion

    #region Constructor
    public Utf8Prober()
    {
        _numOfMbChar = 0;
        _codingSm = new CodingStateMachine(new UTF8SMModel());
        Reset();
    }
    #endregion

    #region GetCharsetName
    public override string GetCharsetName()
    {
        return "UTF-8";
    }
    #endregion

    #region Reset
    public sealed override void Reset()
    {
        _codingSm.Reset();
        _numOfMbChar = 0;
        State = ProbingState.Detecting;
    }
    #endregion

    #region HandleData
    public override ProbingState HandleData(byte[] buf, int offset, int len)
    {
        var max = offset + len;

        for (var i = offset; i < max; i++)
        {
            var codingState = _codingSm.NextState(buf[i]);

            if (codingState == SmModel.Error)
            {
                State = ProbingState.NotMe;
                break;
            }

            if (codingState == SmModel.ItsMe)
            {
                State = ProbingState.FoundIt;
                break;
            }

            if (codingState != SmModel.Start) continue;
            if (_codingSm.CurrentCharLen >= 2)
                _numOfMbChar++;
        }

        if (State != ProbingState.Detecting) return State;
        if (GetConfidence() > ShortcutThreshold)
            State = ProbingState.FoundIt;
        
        return State;
    }
    #endregion

    #region GetConfidence
    public override float GetConfidence()
    {
        var unlike = 0.99f;
        float confidence;

        if (_numOfMbChar < 6)
        {
            for (var i = 0; i < _numOfMbChar; i++)
                unlike *= OneCharProb;
            confidence = 1.0f - unlike;
        }
        else
        {
            confidence = 0.99f;
        }

        return confidence;
    }
    #endregion
}