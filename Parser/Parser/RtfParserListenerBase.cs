// name       : RtfParserListenerBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf.Parser
{
    public class RtfParserListenerBase : IRtfParserListener
    {
        // Members

        public int Level { get; private set; } // Level

        public void ParseBegin()
        {
            Level = 0; // in case something interrupted the normal flow of things previously ...
            DoParseBegin();
        } // ParseBegin

        public void GroupBegin()
        {
            DoGroupBegin();
            Level++;
        } // GroupBegin

        public void TagFound(IRtfTag tag)
        {
            if (tag != null)
                DoTagFound(tag);
        } // TagFound

        public void TextFound(IRtfText text)
        {
            if (text != null)
                DoTextFound(text);
        } // TextFound

        public void GroupEnd()
        {
            Level--;
            DoGroupEnd();
        } // GroupEnd

        public void ParseSuccess()
        {
            DoParseSuccess();
        } // ParseSuccess

        public void ParseFail(RtfException reason)
        {
            DoParseFail(reason);
        } // ParseFail

        public void ParseEnd()
        {
            DoParseEnd();
        } // ParseEnd

        protected virtual void DoParseBegin()
        {
        } // DoParseBegin

        protected virtual void DoGroupBegin()
        {
        } // DoGroupBegin

        protected virtual void DoTagFound(IRtfTag tag)
        {
        } // DoTagFound

        protected virtual void DoTextFound(IRtfText text)
        {
        } // DoTextFound

        protected virtual void DoGroupEnd()
        {
        } // DoGroupEnd

        protected virtual void DoParseSuccess()
        {
        } // DoParseSuccess

        protected virtual void DoParseFail(RtfException reason)
        {
        } // DoParseFail

        protected virtual void DoParseEnd()
        {
        } // DoParseEnd
    } // class RtfParserListenerBase
}