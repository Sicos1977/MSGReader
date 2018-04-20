// name       : RtfTextBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Text;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    public sealed class RtfTextBuilder : RtfElementVisitorBase
    {
        // Members
        private readonly StringBuilder _buffer = new StringBuilder();

        public string CombinedText
        {
            get { return _buffer.ToString(); }
        } // CombinedText

        public RtfTextBuilder() :
            base(RtfElementVisitorOrder.DepthFirst)
        {
            Reset();
        } // RtfTextBuilder

        public void Reset()
        {
            _buffer.Remove(0, _buffer.Length);
        } // Reset

        protected override void DoVisitText(IRtfText text)
        {
            _buffer.Append(text.Text);
        } // DoVisitText
    } // class RtfTextBuilder
}