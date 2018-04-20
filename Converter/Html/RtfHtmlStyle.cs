// name       : RtfHtmlStyle.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.02
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using Itenso.Sys;

namespace Itenso.Rtf.Converter.Html
{
    public class RtfHtmlStyle : IRtfHtmlStyle
    {
        public static RtfHtmlStyle Empty = new RtfHtmlStyle();

        // Members

        public string ForegroundColor { get; set; } // ForegroundColor

        public string BackgroundColor { get; set; } // BackgroundColor

        public string FontFamily { get; set; } // FontFamily

        public string FontSize { get; set; } // FontSize

        public bool IsEmpty => Equals(Empty);
        // IsEmpty

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

        private bool IsEqual(object obj)
        {
            var compare = obj as RtfHtmlStyle; // guaranteed to be non-null
            return
                compare != null &&
                string.Equals(ForegroundColor, compare.ForegroundColor) &&
                string.Equals(BackgroundColor, compare.BackgroundColor) &&
                string.Equals(FontFamily, compare.FontFamily) &&
                string.Equals(FontSize, compare.FontSize);
        } // IsEqual

        private int ComputeHashCode()
        {
            var hash = ForegroundColor.GetHashCode();
            hash = HashTool.AddHashCode(hash, BackgroundColor);
            hash = HashTool.AddHashCode(hash, FontFamily);
            hash = HashTool.AddHashCode(hash, FontSize);
            return hash;
        } // ComputeHashCode
    }
}