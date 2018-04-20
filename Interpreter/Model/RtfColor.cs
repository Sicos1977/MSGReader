// -- FILE ------------------------------------------------------------------
// name       : RtfColor.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Drawing;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    // ------------------------------------------------------------------------
    public sealed class RtfColor : IRtfColor
    {
        // ----------------------------------------------------------------------
        public static readonly IRtfColor Black = new RtfColor(0, 0, 0);
        public static readonly IRtfColor White = new RtfColor(255, 255, 255);

        // ----------------------------------------------------------------------
        // members

        // ----------------------------------------------------------------------
        public RtfColor(int red, int green, int blue)
        {
            if (red < 0 || red > 255)
                throw new RtfColorException(Strings.InvalidColor(red));
            if (green < 0 || green > 255)
                throw new RtfColorException(Strings.InvalidColor(green));
            if (blue < 0 || blue > 255)
                throw new RtfColorException(Strings.InvalidColor(blue));
            Red = red;
            Green = green;
            Blue = blue;
            AsDrawingColor = Color.FromArgb(red, green, blue);
        } // RtfColor

        // ----------------------------------------------------------------------
        public int Red { get; } // Red

        // ----------------------------------------------------------------------
        public int Green { get; } // Green

        // ----------------------------------------------------------------------
        public int Blue { get; } // Blue

        // ----------------------------------------------------------------------
        public Color AsDrawingColor { get; } // AsDrawingColor

        // ----------------------------------------------------------------------
        public override bool Equals(object obj)
        {
            if (obj == this)
                return true;

            if (obj == null || GetType() != obj.GetType())
                return false;

            return IsEqual(obj);
        } // Equals

        // ----------------------------------------------------------------------
        public override int GetHashCode()
        {
            return HashTool.AddHashCode(GetType().GetHashCode(), ComputeHashCode());
        } // GetHashCode

        // ----------------------------------------------------------------------
        public override string ToString()
        {
            return "Color{" + Red + "," + Green + "," + Blue + "}";
        } // ToString

        // ----------------------------------------------------------------------
        private bool IsEqual(object obj)
        {
            var compare = obj as RtfColor; // guaranteed to be non-null
            return compare != null && Red == compare.Red &&
                   Green == compare.Green &&
                   Blue == compare.Blue;
        } // IsEqual

        // ----------------------------------------------------------------------
        private int ComputeHashCode()
        {
            var hash = Red;
            hash = HashTool.AddHashCode(hash, Green);
            hash = HashTool.AddHashCode(hash, Blue);
            return hash;
        } // ComputeHashCode
    } // class RtfColor
} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------