// -- FILE ------------------------------------------------------------------
// name       : RtfTextConverter.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using System.Globalization;
using System.Text;
using Itenso.Rtf.Interpreter;

namespace Itenso.Rtf.Converter.Text
{
    // ------------------------------------------------------------------------
    public class RtfTextConverter : RtfInterpreterListenerBase
    {
        // ----------------------------------------------------------------------
        public const string DefaultTextFileExtension = ".txt";

        // ----------------------------------------------------------------------
        // members
        private readonly StringBuilder plainText = new StringBuilder();

        // ----------------------------------------------------------------------
        public string PlainText
        {
            get { return plainText.ToString(); }
        } // PlainText

        // ----------------------------------------------------------------------
        public RtfTextConvertSettings Settings { get; } // Settings

        // ----------------------------------------------------------------------
        public RtfTextConverter() :
            this(new RtfTextConvertSettings())
        {
        } // RtfTextConverter

        // ----------------------------------------------------------------------
        public RtfTextConverter(RtfTextConvertSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException("settings");

            Settings = settings;
        } // RtfTextConverter

        // ----------------------------------------------------------------------
        public void Clear()
        {
            plainText.Remove(0, plainText.Length);
        } // Clear

        // ----------------------------------------------------------------------
        protected override void DoBeginDocument(IRtfInterpreterContext context)
        {
            Clear();
        } // DoBeginDocument

        // ----------------------------------------------------------------------
        protected override void DoInsertText(IRtfInterpreterContext context, string text)
        {
            if (context.CurrentTextFormat == null)
                return;
            if (!context.CurrentTextFormat.IsHidden || Settings.IsShowHiddenText)
                plainText.Append(text);
        } // DoInsertText

        // ----------------------------------------------------------------------
        protected override void DoInsertSpecialChar(IRtfInterpreterContext context, RtfVisualSpecialCharKind kind)
        {
            switch (kind)
            {
                case RtfVisualSpecialCharKind.Tabulator:
                    plainText.Append(Settings.TabulatorText);
                    break;
                case RtfVisualSpecialCharKind.NonBreakingSpace:
                    plainText.Append(Settings.NonBreakingSpaceText);
                    break;
                case RtfVisualSpecialCharKind.EmSpace:
                    plainText.Append(Settings.EmSpaceText);
                    break;
                case RtfVisualSpecialCharKind.EnSpace:
                    plainText.Append(Settings.EnSpaceText);
                    break;
                case RtfVisualSpecialCharKind.QmSpace:
                    plainText.Append(Settings.QmSpaceText);
                    break;
                case RtfVisualSpecialCharKind.EmDash:
                    plainText.Append(Settings.EmDashText);
                    break;
                case RtfVisualSpecialCharKind.EnDash:
                    plainText.Append(Settings.EnDashText);
                    break;
                case RtfVisualSpecialCharKind.OptionalHyphen:
                    plainText.Append(Settings.OptionalHyphenText);
                    break;
                case RtfVisualSpecialCharKind.NonBreakingHyphen:
                    plainText.Append(Settings.NonBreakingHyphenText);
                    break;
                case RtfVisualSpecialCharKind.Bullet:
                    plainText.Append(Settings.BulletText);
                    break;
                case RtfVisualSpecialCharKind.LeftSingleQuote:
                    plainText.Append(Settings.LeftSingleQuoteText);
                    break;
                case RtfVisualSpecialCharKind.RightSingleQuote:
                    plainText.Append(Settings.RightSingleQuoteText);
                    break;
                case RtfVisualSpecialCharKind.LeftDoubleQuote:
                    plainText.Append(Settings.LeftDoubleQuoteText);
                    break;
                case RtfVisualSpecialCharKind.RightDoubleQuote:
                    plainText.Append(Settings.RightDoubleQuoteText);
                    break;
                default:
                    plainText.Append(Settings.UnknownSpecialCharText);
                    break;
            }
        } // DoInsertSpecialChar

        // ----------------------------------------------------------------------
        protected override void DoInsertBreak(IRtfInterpreterContext context, RtfVisualBreakKind kind)
        {
            switch (kind)
            {
                case RtfVisualBreakKind.Line:
                    plainText.Append(Settings.LineBreakText);
                    break;
                case RtfVisualBreakKind.Page:
                    plainText.Append(Settings.PageBreakText);
                    break;
                case RtfVisualBreakKind.Paragraph:
                    plainText.Append(Settings.ParagraphBreakText);
                    break;
                case RtfVisualBreakKind.Section:
                    plainText.Append(Settings.SectionBreakText);
                    break;
                default:
                    plainText.Append(Settings.UnknownBreakText);
                    break;
            }
        } // DoInsertBreak

        // ----------------------------------------------------------------------
        protected override void DoInsertImage(IRtfInterpreterContext context,
            RtfVisualImageFormat format,
            int width, int height, int desiredWidth, int desiredHeight,
            int scaleWidthPercent, int scaleHeightPercent,
            string imageDataHex
        )
        {
            var imageFormatText = Settings.ImageFormatText;
            if (string.IsNullOrEmpty(imageFormatText))
                return;

            var imageText = string.Format(
                CultureInfo.InvariantCulture,
                imageFormatText,
                format,
                width,
                height,
                desiredWidth,
                desiredHeight,
                scaleWidthPercent,
                scaleHeightPercent,
                imageDataHex);

            plainText.Append(imageText);
        } // DoInsertImage
    } // class RtfTextConverter
} // namespace Itenso.Rtf.Converter.Text
// -- EOF -------------------------------------------------------------------