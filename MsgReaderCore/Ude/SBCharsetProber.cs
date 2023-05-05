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

internal class SingleByteCharSetProber : CharsetProber
{
    #region Consts
    private const int PR_SAMPLE_SIZE = 64;
    private const int PR_SB_ENOUGH_REL_THRESHOLD = 1024;
    private const float PR_POSITIVE_SHORTCUT_THRESHOLD = 0.95f;
    private const float PR_NEGATIVE_SHORTCUT_THRESHOLD = 0.05f;
    private const int PR_SYMBOL_CAT_ORDER = 250;
    private const int PR_NUMBER_OF_SEQ_CAT = 4;
    private const int PR_POSITIVE_CAT = PR_NUMBER_OF_SEQ_CAT - 1;
    #endregion

    #region Fields
    protected SequenceModel Model;

    // true if we need to reverse every pair in the model lookup        
    private readonly bool _reversed;

    // char order of last character
    private byte _lastOrder;

    private int _totalSeqs;
    private int _totalChar;
    private readonly int[] _seqCounters = new int[PR_NUMBER_OF_SEQ_CAT];

    // characters that fall in our sampling range
    private int _freqChar;

    // Optional auxiliary prober for name decision. created and destroyed by the GroupProber
    private readonly CharsetProber _nameProber;
    #endregion

    #region Constructors
    internal SingleByteCharSetProber(SequenceModel model) : this(model, false, null)
    {
    }

    public SingleByteCharSetProber(SequenceModel model, bool reversed, CharsetProber nameProber)
    {
        Model = model;
        _reversed = reversed;
        _nameProber = nameProber;
        Reset();
    }
    #endregion

    #region HandleData
    public override ProbingState HandleData(byte[] buf, int offset, int len)
    {
        var max = offset + len;

        for (var i = offset; i < max; i++)
        {
            var order = Model.GetOrder(buf[i]);

            if (order < PR_SYMBOL_CAT_ORDER)
                _totalChar++;

            if (order < PR_SAMPLE_SIZE)
            {
                _freqChar++;

                if (_lastOrder < PR_SAMPLE_SIZE)
                {
                    _totalSeqs++;
                    if (!_reversed)
                        ++_seqCounters[Model.GetPrecedence(_lastOrder * PR_SAMPLE_SIZE + order)];
                    else // reverse the order of the letters in the lookup
                        ++_seqCounters[Model.GetPrecedence(order * PR_SAMPLE_SIZE + _lastOrder)];
                }
            }

            _lastOrder = order;
        }

        if (State != ProbingState.Detecting) return State;
        if (_totalSeqs <= PR_SB_ENOUGH_REL_THRESHOLD) return State;
        var cf = GetConfidence();
        State = cf switch
        {
            > PR_POSITIVE_SHORTCUT_THRESHOLD => ProbingState.FoundIt,
            < PR_NEGATIVE_SHORTCUT_THRESHOLD => ProbingState.NotMe,
            _ => State
        };

        return State;
    }
    #endregion

    #region DumpStatus
    public override string DumpStatus()
    {
        return $"SBCS: {GetConfidence()} [{GetCharsetName()}]";
    }
    #endregion

    #region GetConfidence
    public override float GetConfidence()
    {
        /*
        NEGATIVE_APPROACH
        if (totalSeqs > 0) {
            if (totalSeqs > seqCounters[NEGATIVE_CAT] * 10)
                return (totalSeqs - seqCounters[NEGATIVE_CAT] * 10)/totalSeqs * freqChar / mTotalChar;
        }
        return 0.01f;
        */
        // POSITIVE_APPROACH

        if (_totalSeqs <= 0) return 0.01f;
        var r = 1.0f * _seqCounters[PR_POSITIVE_CAT] / _totalSeqs / Model.TypicalPositiveRatio;
        r = r * _freqChar / _totalChar;
        if (r >= 1.0f)
            r = 0.99f;
        return r;
    }
    #endregion

    #region Reset
    public sealed override void Reset()
    {
        State = ProbingState.Detecting;
        _lastOrder = 255;
        for (var i = 0; i < PR_NUMBER_OF_SEQ_CAT; i++)
            _seqCounters[i] = 0;
        _totalSeqs = 0;
        _totalChar = 0;
        _freqChar = 0;
    }
    #endregion

    #region GetCharsetName
    public override string GetCharsetName()
    {
        return _nameProber == null
            ? Model.CharsetName
            : _nameProber.GetCharsetName();
    }
    #endregion
}