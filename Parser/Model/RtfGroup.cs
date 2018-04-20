// name       : RtfGroup.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public sealed class RtfGroup : RtfElement, IRtfGroup
    {
        // members

        public RtfElementCollection WritableContents { get; } = new RtfElementCollection();

// WritableContents

        public RtfGroup() :
            base(RtfElementKind.Group)
        {
        } // RtfGroup

        public IRtfElementCollection Contents
        {
            get { return WritableContents; }
        } // Contents

        public string Destination
        {
            get
            {
                if (WritableContents.Count > 0)
                {
                    var firstChild = WritableContents[0];
                    if (firstChild.Kind == RtfElementKind.Tag)
                    {
                        var firstTag = (IRtfTag) firstChild;
                        if (RtfSpec.TagExtensionDestination.Equals(firstTag.Name))
                            if (WritableContents.Count > 1)
                            {
                                var secondChild = WritableContents[1];
                                if (secondChild.Kind == RtfElementKind.Tag)
                                {
                                    var secondTag = (IRtfTag) secondChild;
                                    return secondTag.Name;
                                }
                            }
                        return firstTag.Name;
                    }
                }
                return null;
            }
        } // Destination

        public bool IsExtensionDestination
        {
            get
            {
                if (WritableContents.Count > 0)
                {
                    var firstChild = WritableContents[0];
                    if (firstChild.Kind == RtfElementKind.Tag)
                    {
                        var firstTag = (IRtfTag) firstChild;
                        if (RtfSpec.TagExtensionDestination.Equals(firstTag.Name))
                            return true;
                    }
                }
                return false;
            }
        } // IsExtensionDestination

        public IRtfGroup SelectChildGroupWithDestination(string destination)
        {
            if (destination == null)
                throw new ArgumentNullException("destination");
            foreach (IRtfElement child in WritableContents)
                if (child.Kind == RtfElementKind.Group)
                {
                    var group = (IRtfGroup) child;
                    if (destination.Equals(group.Destination))
                        return group;
                }
            return null;
        } // SelectChildGroupWithDestination

        public override string ToString()
        {
            var buf = new StringBuilder("{");
            var count = WritableContents.Count;
            buf.Append(count);
            buf.Append(" items");
            if (count > 0)
            {
                // visualize the first two child elements for convenience during debugging
                buf.Append(": [");
                buf.Append(WritableContents[0]);
                if (count > 1)
                {
                    buf.Append(", ");
                    buf.Append(WritableContents[1]);
                    if (count > 2)
                    {
                        buf.Append(", ");
                        if (count > 3)
                            buf.Append("..., ");
                        buf.Append(WritableContents[count - 1]);
                    }
                }
                buf.Append("]");
            }
            buf.Append("}");
            return buf.ToString();
        } // ToString

        protected override void DoVisit(IRtfElementVisitor visitor)
        {
            visitor.VisitGroup(this);
        } // DoVisit

        protected override bool IsEqual(object obj)
        {
            var compare = obj as RtfGroup; // guaranteed to be non-null
            return compare != null && base.IsEqual(obj) &&
                   WritableContents.Equals(compare.WritableContents);
        } // IsEqual

        protected override int ComputeHashCode()
        {
            return HashTool.AddHashCode(base.ComputeHashCode(), WritableContents);
        } // ComputeHashCode
    } // class RtfGroup
}