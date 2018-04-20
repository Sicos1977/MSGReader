// name       : RtfFont.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public sealed class RtfFont : IRtfFont
    {
        private readonly int _codePage;

        // Members

        public RtfFont(string id, RtfFontKind kind, RtfFontPitch pitch, int charSet, int codePage, string name)
        {
            if (id == null)
                throw new ArgumentNullException(nameof(id));
            if (charSet < 0)
                throw new ArgumentException(Strings.InvalidCharacterSet(charSet));
            if (codePage < 0)
                throw new ArgumentException(Strings.InvalidCodePage(codePage));
            if (name == null)
                throw new ArgumentNullException(nameof(name));
            Id = id;
            Kind = kind;
            Pitch = pitch;
            CharSet = charSet;
            _codePage = codePage;
            Name = name;
        } // RtfFont

        public string Id { get; } // Id

        public RtfFontKind Kind { get; } // Kind

        public RtfFontPitch Pitch { get; } // Pitch

        public int CharSet { get; } // CharSet

        public int CodePage
        {
            get
            {
                // if a codepage is specified, it overrides the charset setting
                if (_codePage == 0)
                    return RtfSpec.GetCodePage(CharSet);
                return _codePage;
            }
        } // CodePage

        public string Name { get; } // Name

        public Encoding GetEncoding()
        {
            return Encoding.GetEncoding(CodePage);
        } // GetEncoding

        public override bool Equals(object obj)
        {
            if (obj == this)
                return true;

            if (obj == null || GetType() != obj.GetType())
                return false;

            return IsEqual(obj);
        } // Equals

        public override int GetHashCode()
        {
            return HashTool.AddHashCode(GetType().GetHashCode(), ComputeHashCode());
        } // GetHashCode

        public override string ToString()
        {
            return Id + ":" + Name;
        } // ToString

        private bool IsEqual(object obj)
        {
            var compare = obj as RtfFont; // guaranteed to be non-null
            return
                compare != null &&
                Id.Equals(compare.Id) &&
                Kind == compare.Kind &&
                Pitch == compare.Pitch &&
                CharSet == compare.CharSet &&
                _codePage == compare._codePage &&
                Name.Equals(compare.Name);
        } // IsEqual

        private int ComputeHashCode()
        {
            var hash = Id.GetHashCode();
            hash = HashTool.AddHashCode(hash, Kind);
            hash = HashTool.AddHashCode(hash, Pitch);
            hash = HashTool.AddHashCode(hash, CharSet);
            hash = HashTool.AddHashCode(hash, _codePage);
            hash = HashTool.AddHashCode(hash, Name);
            return hash;
        } // ComputeHashCode
    } // class RtfFont
}