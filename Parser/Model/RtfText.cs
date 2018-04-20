// name       : RtfText.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public sealed class RtfText : RtfElement, IRtfText
    {
        // Members

        public RtfText(string text) :
            base(RtfElementKind.Text)
        {
            if (text == null)
                throw new ArgumentNullException(nameof(text));
            Text = text;
        } // RtfText

        public string Text { get; } // Text

        public override string ToString()
        {
            return Text;
        } // ToString

        protected override void DoVisit(IRtfElementVisitor visitor)
        {
            visitor.VisitText(this);
        } // DoVisit

        protected override bool IsEqual(object obj)
        {
            var compare = obj as RtfText; // guaranteed to be non-null
            return compare != null && base.IsEqual(obj) &&
                   Text.Equals(compare.Text);
        } // IsEqual

        protected override int ComputeHashCode()
        {
            return HashTool.AddHashCode(base.ComputeHashCode(), Text);
        } // ComputeHashCode
    } // class RtfText
}