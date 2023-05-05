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

internal class EucjpProber : CharsetProber
{
    #region Fields
    private readonly CodingStateMachine _codingSm;
    private readonly EucjpContextAnalyser _contextAnalyzer;
    private readonly EucjpDistributionAnalyzer _distributionAnalyzer;
    private readonly byte[] _lastChar = new byte[2];
    #endregion

    #region Constructor
    internal EucjpProber()
    {
        _codingSm = new CodingStateMachine(new EUCJPSMModel());
        _distributionAnalyzer = new EucjpDistributionAnalyzer();
        _contextAnalyzer = new EucjpContextAnalyser();
        Reset();
    }
    #endregion

    #region GetCharsetName
    public override string GetCharsetName()
    {
        return "EUC-JP";
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
            var charLen = _codingSm.CurrentCharLen;
            if (i == offset)
            {
                _lastChar[1] = buf[offset];
                _contextAnalyzer.HandleOneChar(_lastChar, 0, charLen);
                _distributionAnalyzer.HandleOneChar(_lastChar, 0, charLen);
            }
            else
            {
                _contextAnalyzer.HandleOneChar(buf, i - 1, charLen);
                _distributionAnalyzer.HandleOneChar(buf, i - 1, charLen);
            }
        }

        _lastChar[0] = buf[max - 1];
        if (State != ProbingState.Detecting) return State;
        if (_contextAnalyzer.GotEnoughData() && GetConfidence() > ShortcutThreshold)
            State = ProbingState.FoundIt;

        return State;
    }
    #endregion

    #region Reset
    public sealed override void Reset()
    {
        _codingSm.Reset();
        State = ProbingState.Detecting;
        _contextAnalyzer.Reset();
        _distributionAnalyzer.Reset();
    }
    #endregion

    #region GetConfidence
    public override float GetConfidence()
    {
        var contextCf = _contextAnalyzer.GetConfidence();
        var distributionCf = _distributionAnalyzer.GetConfidence();
        return contextCf > distributionCf ? contextCf : distributionCf;
    }
    #endregion
}