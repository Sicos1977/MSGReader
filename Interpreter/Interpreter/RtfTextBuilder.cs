// -- FILE ------------------------------------------------------------------
// name       : RtfTextBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Text;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{
    // ------------------------------------------------------------------------
    public sealed class RtfTextBuilder : RtfElementVisitorBase
    {
        // ----------------------------------------------------------------------
        // members
        private readonly StringBuilder buffer = new StringBuilder();

        // ----------------------------------------------------------------------
        public string CombinedText
        {
            get { return buffer.ToString(); }
        } // CombinedText

        // ----------------------------------------------------------------------
        public RtfTextBuilder() :
            base(RtfElementVisitorOrder.DepthFirst)
        {
            Reset();
        } // RtfTextBuilder

        // ----------------------------------------------------------------------
        public void Reset()
        {
            buffer.Remove(0, buffer.Length);
        } // Reset

        // ----------------------------------------------------------------------
        protected override void DoVisitText(IRtfText text)
        {
            buffer.Append(text.Text);
        } // DoVisitText
    } // class RtfTextBuilder
} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------