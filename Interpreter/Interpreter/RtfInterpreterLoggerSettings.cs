// name       : RtfInterpreterLoggerSettings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.05.29
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

namespace Itenso.Rtf.Interpreter
{
    public class RtfInterpreterLoggerSettings
    {
        // Members

        public bool Enabled { get; set; }

        public string BeginDocumentText { get; set; } = Strings.LogBeginDocument;

// BeginDocumentText

        public string EndDocumentText { get; set; } = Strings.LogEndDocument;

// EndDocumentText

        public string TextFormatText { get; set; } = Strings.LogInsertText;

// TextFormatText

        public string TextOverflowText { get; set; } = Strings.LogOverflowText;

// TextOverflowText

        public string SpecialCharFormatText { get; set; } = Strings.LogInsertChar;

// SpecialCharFormatText

        public string BreakFormatText { get; set; } = Strings.LogInsertBreak;

// BreakFormatText

        public string ImageFormatText { get; set; } = Strings.LogInsertImage;

// ImageFormatText

        public int TextMaxLength { get; set; } = 80;

// TextMaxLength

        public RtfInterpreterLoggerSettings() :
            this(true)
        {
        } // RtfInterpreterLoggerSettings

        public RtfInterpreterLoggerSettings(bool enabled)
        {
            Enabled = enabled;
        } // RtfInterpreterLoggerSettings
    } // class RtfInterpreterLoggerSettings
}