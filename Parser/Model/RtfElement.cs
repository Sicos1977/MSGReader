// name       : RtfElement.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public abstract class RtfElement : IRtfElement
    {
        // Members

        protected RtfElement(RtfElementKind kind)
        {
            Kind = kind;
        } // RtfElement

        public RtfElementKind Kind { get; } // Kind

        public void Visit(IRtfElementVisitor visitor)
        {
            if (visitor == null)
                throw new ArgumentNullException(nameof(visitor));
            DoVisit(visitor);
        } // Visit

        public sealed override bool Equals(object obj)
        {
            if (obj == this)
                return true;

            if (obj == null || GetType() != obj.GetType())
                return false;

            return IsEqual(obj);
        } // Equals

        public sealed override int GetHashCode()
        {
            return HashTool.AddHashCode(GetType().GetHashCode(), ComputeHashCode());
        } // GetHashCode

        protected abstract void DoVisit(IRtfElementVisitor visitor);

        protected virtual bool IsEqual(object obj)
        {
            return true;
        } // IsEqual

        protected virtual int ComputeHashCode()
        {
            return 0x0f00ba11;
        } // ComputeHashCode
    } // class RtfElement
}