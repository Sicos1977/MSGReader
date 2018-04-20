// name       : RtfTextConvertSettings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.05.29
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Rtf.Converter.Text
{
    public class RtfTextConvertSettings
    {
        public const string SeparatorCr = "\r";
        public const string SeparatorLf = "\n";
        public const string SeparatorCrLf = "\r\n";
        public const string SeparatorLfCr = "\n\r";

        // members: image

        // members: breaks

        // members: hidden text

        // members: special chars

        public bool IsShowHiddenText { get; set; }

        public string TabulatorText { get; set; } = "\t";

// TabulatorText

        public string NonBreakingSpaceText { get; set; } = " ";

// NonBreakingSpaceText

        public string EmSpaceText { get; set; } = " ";

// EmSpaceText

        public string EnSpaceText { get; set; } = " ";

// EnSpaceText

        public string QmSpaceText { get; set; } = " ";

// QmSpaceText

        public string EmDashText { get; set; } = "-";

// EmDashText

        public string EnDashText { get; set; } = "-";

// EnDashText

        public string OptionalHyphenText { get; set; } = "-";

// OptionalHyphenText

        public string NonBreakingHyphenText { get; set; } = "-";

// NonBreakingHyphenText

        public string BulletText { get; set; } = "°";

// BulletText

        public string LeftSingleQuoteText { get; set; } = "`";

// LeftSingleQuoteText

        public string RightSingleQuoteText { get; set; } = "´";

// RightSingleQuoteText

        public string LeftDoubleQuoteText { get; set; } = "``";

// LeftDoubleQuoteText

        public string RightDoubleQuoteText { get; set; } = "´´";

// RightDoubleQuoteText

        public string UnknownSpecialCharText { get; set; }

        public string LineBreakText { get; set; } // LineBreakText

        public string PageBreakText { get; set; } // PageBreakText

        public string ParagraphBreakText { get; set; } // ParagraphBreakText

        public string SectionBreakText { get; set; } // SectionBreakText

        public string UnknownBreakText { get; set; } // UnknownBreakText

        public string ImageFormatText { get; set; } = Strings.ImageFormatText;

// ImageFormatText

        public RtfTextConvertSettings() :
            this(SeparatorCrLf)
        {
        } // RtfTextConvertSettings

        public RtfTextConvertSettings(string breakText)
        {
            SetBreakText(breakText);
        } // RtfTextConvertSettings

        public void SetBreakText(string breakText)
        {
            if (breakText == null)
                throw new ArgumentNullException("breakText");

            LineBreakText = breakText;
            PageBreakText = breakText + breakText;
            ParagraphBreakText = breakText;
            SectionBreakText = breakText + breakText;
            UnknownBreakText = breakText;
        } // SetBreakText
    } // class RtfTextConvertSettings
}