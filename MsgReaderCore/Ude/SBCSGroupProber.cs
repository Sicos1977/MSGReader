//
// Version: MPL 1.1/GPL 2.0/LGPL 2.1
//
// The contents of this file are subject to the Mozilla Public License Version
// 1.1 (the "License"); you may not use this file except in compliance with
// the License. You may obtain a copy of the License at
// http://www.mozilla.org/MPL/
//
// Software distributed under the License is distributed on an "AS IS" basis,
// WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
// for the specific language governing rights and limitations under the
// License.
//
// The Original Code is Mozilla Universal charset detector code.
//
// The Initial Developer of the Original Code is
// Netscape Communications Corporation.
// Portions created by the Initial Developer are Copyright (C) 2001
// the Initial Developer. All Rights Reserved.
//
// Contributor(s):
//          Shy Shalom <shooshX@gmail.com>
//          Rudi Pettazzi <rudi.pettazzi@gmail.com> (C# port)
// 
// Alternatively, the contents of this file may be used under the terms of
// either the GNU General Public License Version 2 or later (the "GPL"), or
// the GNU Lesser General Public License Version 2.1 or later (the "LGPL"),
// in which case the provisions of the GPL or the LGPL are applicable instead
// of those above. If you wish to allow use of your version of this file only
// under the terms of either the GPL or the LGPL, and not to allow others to
// use your version of this file under the terms of the MPL, indicate your
// decision by deleting the provisions above and replace them with the notice
// and other provisions required by the GPL or the LGPL. If you do not delete
// the provisions above, a recipient may use your version of this file under
// the terms of any one of the MPL, the GPL or the LGPL.
//

using System.Text;

namespace MsgReader.Ude;

internal class SbcsGroupProber : CharsetProber
{
    #region Consts
    private const int PR_PROBERS_NUM = 13;
    #endregion

    #region Fields
    private readonly CharsetProber[] _probers = new CharsetProber[PR_PROBERS_NUM];
    private readonly bool[] _isActive = new bool[PR_PROBERS_NUM];
    private int _bestGuess;
    private int _activeNum;
    #endregion

    #region Constructor
    internal SbcsGroupProber()
    {
        _probers[0] = new SingleByteCharSetProber(new Win1251Model());
        _probers[1] = new SingleByteCharSetProber(new Koi8RModel());
        _probers[2] = new SingleByteCharSetProber(new Latin5Model());
        _probers[3] = new SingleByteCharSetProber(new MacCyrillicModel());
        _probers[4] = new SingleByteCharSetProber(new Ibm866Model());
        _probers[5] = new SingleByteCharSetProber(new Ibm855Model());
        _probers[6] = new SingleByteCharSetProber(new Latin7Model());
        _probers[7] = new SingleByteCharSetProber(new Win1253Model());
        _probers[8] = new SingleByteCharSetProber(new Latin5BulgarianModel());
        _probers[9] = new SingleByteCharSetProber(new Win1251BulgarianModel());
        var hebrewProber = new HebrewProber();
        _probers[10] = hebrewProber;
        // Logical  
        _probers[11] = new SingleByteCharSetProber(new Win1255Model(), false, hebrewProber);
        // Visual
        _probers[12] = new SingleByteCharSetProber(new Win1255Model(), true, hebrewProber);
        hebrewProber.SetModelProbers(_probers[11], _probers[12]);
        // disable latin2 before latin1 is available, otherwise all latin1 
        // will be detected as latin2 because of their similarity.
        //probers[13] = new SingleByteCharSetProber(new Latin2HungarianModel());
        //probers[14] = new SingleByteCharSetProber(new Win1250HungarianModel());            
        Reset();
    }
    #endregion

    #region HandleData
    public override ProbingState HandleData(byte[] buf, int offset, int len)
    {
        //apply filter to original buffer, and we got new buffer back
        //depend on what script it is, we will feed them the new buffer 
        //we got after applying proper filter
        //this is done without any consideration to KeepEnglishLetters
        //of each prober since as of now, there are no probers here which
        //recognize languages with English characters.
        var newBuf = FilterWithoutEnglishLetters(buf, offset, len);
        if (newBuf.Length == 0)
            return State; // Nothing to see here, move on.

        for (var i = 0; i < PR_PROBERS_NUM; i++)
        {
            if (!_isActive[i])
                continue;
            var st = _probers[i].HandleData(newBuf, 0, newBuf.Length);

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
                return 0.99f; //sure yes
            case ProbingState.NotMe:
                return 0.01f; //sure no
            case ProbingState.Detecting:
            default:
                for (var i = 0; i < PR_PROBERS_NUM; i++)
                {
                    if (!_isActive[i])
                        continue;
                    var cf = _probers[i].GetConfidence();
                    if (!(bestConf < cf)) continue;
                    bestConf = cf;
                    _bestGuess = i;
                }

                break;
        }

        return bestConf;
    }
    #endregion

    #region DumpStatus
    public override string DumpStatus()
    {
        var sb = new StringBuilder();

        var cf = GetConfidence();
        sb.AppendLine("SBCS Group Prober --------begin status");
        for (var i = 0; i < PR_PROBERS_NUM; i++)
            sb.AppendLine(!_isActive[i]
                ? $" inactive: [{_probers[i].GetCharsetName()}] (i.e. confidence is too low)."
                : _probers[i].DumpStatus());

        var bestMatch = _bestGuess >= 0 ? _probers[_bestGuess].GetCharsetName() : null;

        sb.Append($"SBCS Group found best match [{bestMatch}] confidence {cf}.");

        return sb.ToString();
    }
    #endregion

    #region Reset
    public sealed override void Reset()
    {
        for (var i = 0; i < PR_PROBERS_NUM; i++)
            if (_probers[i] != null)
            {
                _probers[i].Reset();
                _isActive[i] = true;
            }
            else
            {
                _isActive[i] = false;
            }

        _bestGuess = -1;
        State = ProbingState.Detecting;
    }
    #endregion

    #region GetCharsetName
    public override string GetCharsetName()
    {
        //if we have no answer yet
        if (_bestGuess != -1) return _probers[_bestGuess].GetCharsetName();
        GetConfidence();
        //no charset seems positive
        if (_bestGuess == -1)
            _bestGuess = 0;

        return _probers[_bestGuess].GetCharsetName();
    }
    #endregion
}