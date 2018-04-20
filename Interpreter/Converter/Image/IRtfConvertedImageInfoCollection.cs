// name       : IRtfConvertedImageInfoCollection.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.10.13
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Collections;

namespace Itenso.Rtf.Converter.Image
{
    public interface IRtfConvertedImageInfoCollection : IEnumerable
    {
        int Count { get; }

        IRtfConvertedImageInfo this[int index] { get; }

        void CopyTo(IRtfConvertedImageInfo[] array, int index);
    }
}