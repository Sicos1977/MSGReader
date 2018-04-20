// name       : IRtfElementVisitor.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf
{
    public interface IRtfElementVisitor
    {
        void VisitTag(IRtfTag tag);

        void VisitGroup(IRtfGroup group);

        void VisitText(IRtfText text);
    }
}