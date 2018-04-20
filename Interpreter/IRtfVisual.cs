// name       : IRtfVisual.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf
{
    public interface IRtfVisual
    {
        RtfVisualKind Kind { get; }

        void Visit(IRtfVisualVisitor visitor);
    }
}