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

namespace MsgReader.Ude;

internal class Latin1Prober : CharsetProber
{
    #region Consts
    private const int PR_FREQ_CAT_NUM = 4;

    private const int PR_UDF = 0; // undefined
    private const int PR_OTH = 1; // other
    private const int PR_ASC = 2; // ascii capital letter
    private const int PR_ASS = 3; // ascii small letter
    private const int PR_ACV = 4; // accent capital vowel
    private const int PR_ACO = 5; // accent capital other
    private const int PR_ASV = 6; // accent small vowel
    private const int PR_ASO = 7; // accent small other

    private const int PR_CLASS_NUM = 8; // total classes
    #endregion

    #region Fields
    private static readonly byte[] Latin1CharToClass =
    {
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 00 - 07
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 08 - 0F
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 10 - 17
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 18 - 1F
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 20 - 27
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 28 - 2F
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 30 - 37
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 38 - 3F
        PR_OTH, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, // 40 - 47
        PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, // 48 - 4F
        PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, PR_ASC, // 50 - 57
        PR_ASC, PR_ASC, PR_ASC, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 58 - 5F
        PR_OTH, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, // 60 - 67
        PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, // 68 - 6F
        PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, PR_ASS, // 70 - 77
        PR_ASS, PR_ASS, PR_ASS, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 78 - 7F
        PR_OTH, PR_UDF, PR_OTH, PR_ASO, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 80 - 87
        PR_OTH, PR_OTH, PR_ACO, PR_OTH, PR_ACO, PR_UDF, PR_ACO, PR_UDF, // 88 - 8F
        PR_UDF, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // 90 - 97
        PR_OTH, PR_OTH, PR_ASO, PR_OTH, PR_ASO, PR_UDF, PR_ASO, PR_ACO, // 98 - 9F
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // A0 - A7
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // A8 - AF
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // B0 - B7
        PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, PR_OTH, // B8 - BF
        PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACO, PR_ACO, // C0 - C7
        PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, // C8 - CF
        PR_ACO, PR_ACO, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_OTH, // D0 - D7
        PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACV, PR_ACO, PR_ACO, PR_ACO, // D8 - DF
        PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASO, PR_ASO, // E0 - E7
        PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, // E8 - EF
        PR_ASO, PR_ASO, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_OTH, // F0 - F7
        PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASV, PR_ASO, PR_ASO, PR_ASO // F8 - FF
    };

    /* 0 : illegal 
       1 : very unlikely 
       2 : normal 
       3 : very likely
    */
    private static readonly byte[] Latin1ClassModel =
    {
        /*      UDF OTH ASC ASS ACV ACO ASV ASO  */
        /*UDF*/ 0, 0, 0, 0, 0, 0, 0, 0,
        /*OTH*/ 0, 3, 3, 3, 3, 3, 3, 3,
        /*ASC*/ 0, 3, 3, 3, 3, 3, 3, 3,
        /*ASS*/ 0, 3, 3, 3, 1, 1, 3, 3,
        /*ACV*/ 0, 3, 3, 3, 1, 2, 1, 2,
        /*ACO*/ 0, 3, 3, 3, 3, 3, 3, 3,
        /*ASV*/ 0, 3, 1, 3, 1, 1, 1, 3,
        /*ASO*/ 0, 3, 1, 3, 1, 1, 3, 3
    };

    private byte _lastCharClass;
    private readonly int[] _freqCounter = new int[PR_FREQ_CAT_NUM];
    #endregion

    #region Constructor
    internal Latin1Prober()
    {
        Reset();
    }
    #endregion

    #region GetCharsetName
    public override string GetCharsetName()
    {
        return "windows-1252";
    }
    #endregion

    #region Reset
    public sealed override void Reset()
    {
        State = ProbingState.Detecting;
        _lastCharClass = PR_OTH;
        for (var i = 0; i < PR_FREQ_CAT_NUM; i++)
            _freqCounter[i] = 0;
    }
    #endregion

    #region HandleData
    public override ProbingState HandleData(byte[] buf, int offset, int len)
    {
        var newBuffer = FilterWithEnglishLetters(buf, offset, len);

        foreach (var t in newBuffer)
        {
            var charClass = Latin1CharToClass[t];
            var freq = Latin1ClassModel[_lastCharClass * PR_CLASS_NUM + charClass];
            if (freq == 0)
            {
                State = ProbingState.NotMe;
                break;
            }

            _freqCounter[freq]++;
            _lastCharClass = charClass;
        }

        return State;
    }
    #endregion

    #region GetConfidence
    public override float GetConfidence()
    {
        if (State == ProbingState.NotMe)
            return 0.01f;

        float confidence;
        var total = 0;
        for (var i = 0; i < PR_FREQ_CAT_NUM; i++) total += _freqCounter[i];

        if (total <= 0)
        {
            confidence = 0.0f;
        }
        else
        {
            confidence = _freqCounter[3] * 1.0f / total;
            confidence -= _freqCounter[1] * 20.0f / total;
        }

        // lower the confidence of latin1 so that other more accurate detector 
        // can take priority.
        return confidence < 0.0f ? 0.0f : confidence * 0.5f;
    }
    #endregion

    #region DumpStatus
    public override string DumpStatus()
    {
        return $"Latin1Prober: {GetConfidence()} [{GetCharsetName()}]";
    }
    #endregion
}