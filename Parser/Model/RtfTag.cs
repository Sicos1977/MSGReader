// name       : RtfTag.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Globalization;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public sealed class RtfTag : RtfElement, IRtfTag
    {
        // Members

        public RtfTag(string name) :
            base(RtfElementKind.Tag)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));
            FullName = name;
            Name = name;
            ValueAsText = null;
            ValueAsNumber = -1;
        } // RtfTag

        public RtfTag(string name, string value) :
            base(RtfElementKind.Tag)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));
            if (value == null)
                throw new ArgumentNullException(nameof(value));
            FullName = name + value;
            Name = name;
            ValueAsText = value;
            int numericalValue;
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out numericalValue))
                ValueAsNumber = numericalValue;
            else
                ValueAsNumber = -1;
        } // RtfTag

        public string FullName { get; } // FullName

        public string Name { get; } // Name

        public bool HasValue
        {
            get { return ValueAsText != null; }
        } // HasValue

        public string ValueAsText { get; } // ValueAsText

        public int ValueAsNumber { get; } // ValueAsNumber

        public override string ToString()
        {
            return "\\" + FullName;
        } // ToString

        protected override void DoVisit(IRtfElementVisitor visitor)
        {
            visitor.VisitTag(this);
        } // DoVisit

        protected override bool IsEqual(object obj)
        {
            var compare = obj as RtfTag; // guaranteed to be non-null
            return compare != null && base.IsEqual(obj) &&
                   FullName.Equals(compare.FullName);
        } // IsEqual

        protected override int ComputeHashCode()
        {
            var hash = base.ComputeHashCode();
            hash = HashTool.AddHashCode(hash, FullName);
            return hash;
        } // ComputeHashCode
    } // class RtfTag
}