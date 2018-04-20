// name       : RtfVisualText.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.22
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public sealed class RtfVisualText : RtfVisual, IRtfVisualText
    {
        // members
        private IRtfTextFormat format;

        public RtfVisualText(string text, IRtfTextFormat format) :
            base(RtfVisualKind.Text)
        {
            if (text == null)
                throw new ArgumentNullException("text");
            if (format == null)
                throw new ArgumentNullException("format");
            Text = text;
            this.format = format;
        } // RtfVisualText

        public string Text { get; } // Text

        public IRtfTextFormat Format
        {
            get { return format; }
            set
            {
                if (format == null)
                    throw new ArgumentNullException("value");
                format = value;
            }
        } // Format

        protected override void DoVisit(IRtfVisualVisitor visitor)
        {
            visitor.VisitText(this);
        } // DoVisit

        protected override bool IsEqual(object obj)
        {
            var compare = obj as RtfVisualText; // guaranteed to be non-null
            return
                compare != null &&
                base.IsEqual(compare) &&
                Text.Equals(compare.Text) &&
                format.Equals(compare.format);
        } // IsEqual

        protected override int ComputeHashCode()
        {
            var hash = base.ComputeHashCode();
            hash = HashTool.AddHashCode(hash, Text);
            hash = HashTool.AddHashCode(hash, format);
            return hash;
        } // ComputeHashCode

        public override string ToString()
        {
            return "'" + Text + "'";
        } // ToString
    } // class RtfVisualText
}