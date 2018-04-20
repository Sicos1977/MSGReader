// name       : IRtfTextFormatCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Collections;

namespace Itenso.Rtf
{
    public interface IRtfTextFormatCollection : IEnumerable
    {
        int Count { get; }

        IRtfTextFormat this[int index] { get; }

        bool Contains(IRtfTextFormat format);

        int IndexOf(IRtfTextFormat format);

        void CopyTo(IRtfTextFormat[] array, int index);
    }
}