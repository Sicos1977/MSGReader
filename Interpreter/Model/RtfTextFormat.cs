// name       : IRtfTextFormat.cs
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
    public sealed class RtfTextFormat : IRtfTextFormat
    {
        // Members

        public bool IsNormal
        {
            get
            {
                return
                    !IsBold && !IsItalic && !IsUnderline && !IsStrikeThrough &&
                    !IsHidden &&
                    FontSize == RtfSpec.DefaultFontSize &&
                    SuperScript == 0 &&
                    RtfColor.Black.Equals(ForegroundColor) &&
                    RtfColor.White.Equals(BackgroundColor);
            }
        } // IsNormal

        public RtfTextFormat(IRtfFont font, int fontSize)
        {
            if (font == null)
                throw new ArgumentNullException(nameof(font));
            if (fontSize <= 0 || fontSize > 0xFFFF)
                throw new ArgumentException(Strings.FontSizeOutOfRange(fontSize));
            Font = font;
            FontSize = fontSize;
        } // RtfTextFormat

        public RtfTextFormat(IRtfTextFormat copy)
        {
            if (copy == null)
                throw new ArgumentNullException(nameof(copy));
            Font = copy.Font; // enough because immutable
            FontSize = copy.FontSize;
            SuperScript = copy.SuperScript;
            IsBold = copy.IsBold;
            IsItalic = copy.IsItalic;
            IsUnderline = copy.IsUnderline;
            IsStrikeThrough = copy.IsStrikeThrough;
            IsHidden = copy.IsHidden;
            BackgroundColor = copy.BackgroundColor; // enough because immutable
            ForegroundColor = copy.ForegroundColor; // enough because immutable
            Alignment = copy.Alignment;
        } // RtfTextFormat

        public RtfTextFormat(RtfTextFormat copy)
        {
            if (copy == null)
                throw new ArgumentNullException(nameof(copy));
            Font = copy.Font; // enough because immutable
            FontSize = copy.FontSize;
            SuperScript = copy.SuperScript;
            IsBold = copy.IsBold;
            IsItalic = copy.IsItalic;
            IsUnderline = copy.IsUnderline;
            IsStrikeThrough = copy.IsStrikeThrough;
            IsHidden = copy.IsHidden;
            BackgroundColor = copy.BackgroundColor; // enough because immutable
            ForegroundColor = copy.ForegroundColor; // enough because immutable
            Alignment = copy.Alignment;
        } // RtfTextFormat

        public string FontDescriptionDebug
        {
            get
            {
                var buf = new StringBuilder(Font.Name);
                buf.Append(", ");
                buf.Append(FontSize);
                buf.Append(SuperScript >= 0 ? "+" : "");
                buf.Append(SuperScript);
                buf.Append(", ");
                if (IsBold || IsItalic || IsUnderline || IsStrikeThrough)
                {
                    var combined = false;
                    if (IsBold)
                    {
                        buf.Append("bold");
                        combined = true;
                    }
                    if (IsItalic)
                    {
                        buf.Append(combined ? "+italic" : "italic");
                        combined = true;
                    }
                    if (IsUnderline)
                    {
                        buf.Append(combined ? "+underline" : "underline");
                        combined = true;
                    }
                    if (IsStrikeThrough)
                        buf.Append(combined ? "+strikethrough" : "strikethrough");
                }
                else
                {
                    buf.Append("plain");
                }
                if (IsHidden)
                    buf.Append(", hidden");
                return buf.ToString();
            }
        } // FontDescriptionDebug

        public IRtfFont Font { get; private set; } // Font

        public int FontSize { get; private set; } // FontSize

        public int SuperScript { get; private set; } // SuperScript

        public bool IsBold { get; private set; } // IsBold

        public bool IsItalic { get; private set; } // IsItalic

        public bool IsUnderline { get; private set; } // IsUnderline

        public bool IsStrikeThrough { get; private set; } // IsStrikeThrough

        public bool IsHidden { get; private set; } // IsHidden

        public IRtfColor BackgroundColor { get; private set; } = RtfColor.White;

// BackgroundColor

        public IRtfColor ForegroundColor { get; private set; } = RtfColor.Black;

// ForegroundColor

        public RtfTextAlignment Alignment { get; private set; } = RtfTextAlignment.Left;

// Alignment

        IRtfTextFormat IRtfTextFormat.Duplicate()
        {
            return new RtfTextFormat(this);
        } // IRtfTextFormat.Duplicate

        public RtfTextFormat DeriveWithFont(IRtfFont rtfFont)
        {
            if (rtfFont == null)
                throw new ArgumentNullException(nameof(rtfFont));
            if (Font.Equals(rtfFont))
                return this;

            var copy = new RtfTextFormat(this);
            copy.Font = rtfFont;
            return copy;
        } // DeriveWithFont

        public RtfTextFormat DeriveWithFontSize(int derivedFontSize)
        {
            if (derivedFontSize < 0 || derivedFontSize > 0xFFFF)
                throw new ArgumentException(Strings.FontSizeOutOfRange(derivedFontSize));
            if (FontSize == derivedFontSize)
                return this;

            var copy = new RtfTextFormat(this);
            copy.FontSize = derivedFontSize;
            return copy;
        } // DeriveWithFontSize

        public RtfTextFormat DeriveWithSuperScript(int deviation)
        {
            if (SuperScript == deviation)
                return this;

            var copy = new RtfTextFormat(this);
            copy.SuperScript = deviation;
            // reset font size
            if (deviation == 0)
                copy.FontSize = FontSize / 2 * 3;
            return copy;
        } // DeriveWithSuperScript

        public RtfTextFormat DeriveWithSuperScript(bool super)
        {
            var copy = new RtfTextFormat(this);
            copy.FontSize = Math.Max(1, FontSize * 2 / 3);
            copy.SuperScript = (super ? 1 : -1) * Math.Max(1, FontSize / 2);
            return copy;
        } // DeriveWithSuperScript

        public RtfTextFormat DeriveNormal()
        {
            if (IsNormal)
                return this;

            var copy = new RtfTextFormat(Font, RtfSpec.DefaultFontSize);
            copy.Alignment = Alignment; // this is a paragraph property, keep it
            return copy;
        } // DeriveNormal

        public RtfTextFormat DeriveWithBold(bool derivedBold)
        {
            if (IsBold == derivedBold)
                return this;

            var copy = new RtfTextFormat(this);
            copy.IsBold = derivedBold;
            return copy;
        } // DeriveWithBold

        public RtfTextFormat DeriveWithItalic(bool derivedItalic)
        {
            if (IsItalic == derivedItalic)
                return this;

            var copy = new RtfTextFormat(this);
            copy.IsItalic = derivedItalic;
            return copy;
        } // DeriveWithItalic

        public RtfTextFormat DeriveWithUnderline(bool derivedUnderline)
        {
            if (IsUnderline == derivedUnderline)
                return this;

            var copy = new RtfTextFormat(this);
            copy.IsUnderline = derivedUnderline;
            return copy;
        } // DeriveWithUnderline

        public RtfTextFormat DeriveWithStrikeThrough(bool derivedStrikeThrough)
        {
            if (IsStrikeThrough == derivedStrikeThrough)
                return this;

            var copy = new RtfTextFormat(this);
            copy.IsStrikeThrough = derivedStrikeThrough;
            return copy;
        } // DeriveWithStrikeThrough

        public RtfTextFormat DeriveWithHidden(bool derivedHidden)
        {
            if (IsHidden == derivedHidden)
                return this;

            var copy = new RtfTextFormat(this);
            copy.IsHidden = derivedHidden;
            return copy;
        } // DeriveWithHidden

        public RtfTextFormat DeriveWithBackgroundColor(IRtfColor derivedBackgroundColor)
        {
            if (derivedBackgroundColor == null)
                throw new ArgumentNullException(nameof(derivedBackgroundColor));
            if (BackgroundColor.Equals(derivedBackgroundColor))
                return this;

            var copy = new RtfTextFormat(this);
            copy.BackgroundColor = derivedBackgroundColor;
            return copy;
        } // DeriveWithBackgroundColor

        public RtfTextFormat DeriveWithForegroundColor(IRtfColor derivedForegroundColor)
        {
            if (derivedForegroundColor == null)
                throw new ArgumentNullException(nameof(derivedForegroundColor));
            if (ForegroundColor.Equals(derivedForegroundColor))
                return this;

            var copy = new RtfTextFormat(this);
            copy.ForegroundColor = derivedForegroundColor;
            return copy;
        } // DeriveWithForegroundColor

        public RtfTextFormat DeriveWithAlignment(RtfTextAlignment derivedAlignment)
        {
            if (Alignment == derivedAlignment)
                return this;

            var copy = new RtfTextFormat(this);
            copy.Alignment = derivedAlignment;
            return copy;
        } // DeriveWithForegroundColor

        public RtfTextFormat Duplicate()
        {
            return new RtfTextFormat(this);
        } // Duplicate

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

        private bool IsEqual(object obj)
        {
            var compare = obj as RtfTextFormat; // guaranteed to be non-null
            return
                compare != null &&
                Font.Equals(compare.Font) &&
                FontSize == compare.FontSize &&
                SuperScript == compare.SuperScript &&
                IsBold == compare.IsBold &&
                IsItalic == compare.IsItalic &&
                IsUnderline == compare.IsUnderline &&
                IsStrikeThrough == compare.IsStrikeThrough &&
                IsHidden == compare.IsHidden &&
                BackgroundColor.Equals(compare.BackgroundColor) &&
                ForegroundColor.Equals(compare.ForegroundColor) &&
                Alignment == compare.Alignment;
        } // IsEqual

        private int ComputeHashCode()
        {
            var hash = Font.GetHashCode();
            hash = HashTool.AddHashCode(hash, FontSize);
            hash = HashTool.AddHashCode(hash, SuperScript);
            hash = HashTool.AddHashCode(hash, IsBold);
            hash = HashTool.AddHashCode(hash, IsItalic);
            hash = HashTool.AddHashCode(hash, IsUnderline);
            hash = HashTool.AddHashCode(hash, IsStrikeThrough);
            hash = HashTool.AddHashCode(hash, IsHidden);
            hash = HashTool.AddHashCode(hash, BackgroundColor);
            hash = HashTool.AddHashCode(hash, ForegroundColor);
            hash = HashTool.AddHashCode(hash, Alignment);
            return hash;
        } // ComputeHashCode

        public override string ToString()
        {
            var buf = new StringBuilder("Font ");
            buf.Append(FontDescriptionDebug);
            buf.Append(", ");
            buf.Append(Alignment);
            buf.Append(", ");
            buf.Append(ForegroundColor);
            buf.Append(" on ");
            buf.Append(BackgroundColor);
            return buf.ToString();
        } // ToString
    } // class RtfTextFormat
}