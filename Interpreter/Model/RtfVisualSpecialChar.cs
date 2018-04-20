// -- FILE ------------------------------------------------------------------
// name       : RtfVisualSpecialChar.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.22
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    // ------------------------------------------------------------------------
    public sealed class RtfVisualSpecialChar : RtfVisual, IRtfVisualSpecialChar
    {
        // ----------------------------------------------------------------------
        // members

        // ----------------------------------------------------------------------
        public RtfVisualSpecialChar(RtfVisualSpecialCharKind charKind) :
            base(RtfVisualKind.Special)
        {
            CharKind = charKind;
        } // RtfVisualSpecialChar

        // ----------------------------------------------------------------------
        public RtfVisualSpecialCharKind CharKind { get; } // CharKind

        // ----------------------------------------------------------------------
        protected override void DoVisit(IRtfVisualVisitor visitor)
        {
            visitor.VisitSpecial(this);
        } // DoVisit

        // ----------------------------------------------------------------------
        protected override bool IsEqual(object obj)
        {
            var compare = obj as RtfVisualSpecialChar; // guaranteed to be non-null
            return
                compare != null &&
                base.IsEqual(compare) &&
                CharKind == compare.CharKind;
        } // IsEqual

        // ----------------------------------------------------------------------
        protected override int ComputeHashCode()
        {
            return HashTool.AddHashCode(base.ComputeHashCode(), CharKind);
        } // ComputeHashCode

        // ----------------------------------------------------------------------
        public override string ToString()
        {
            return CharKind.ToString();
        } // ToString
    } // class RtfVisualSpecialChar
} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------