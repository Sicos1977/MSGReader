// name       : RtfTextConverter.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Globalization;
using System.Text;
using Itenso.Rtf.Interpreter;

namespace Itenso.Rtf.Converter.Text
{
    public class RtfTextConverter : RtfInterpreterListenerBase
    {
        public const string DefaultTextFileExtension = ".txt";

        // Members
        private readonly StringBuilder _plainText = new StringBuilder();

        public string PlainText
        {
            get { return _plainText.ToString(); }
        } // PlainText

        public RtfTextConvertSettings Settings { get; } // Settings

        public RtfTextConverter() :
            this(new RtfTextConvertSettings())
        {
        } // RtfTextConverter

        public RtfTextConverter(RtfTextConvertSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            Settings = settings;
        } // RtfTextConverter

        public void Clear()
        {
            _plainText.Remove(0, _plainText.Length);
        } // Clear

        protected override void DoBeginDocument(IRtfInterpreterContext context)
        {
            Clear();
        } // DoBeginDocument

        protected override void DoInsertText(IRtfInterpreterContext context, string text)
        {
            if (context.CurrentTextFormat == null)
                return;
            if (!context.CurrentTextFormat.IsHidden || Settings.IsShowHiddenText)
                _plainText.Append(text);
        } // DoInsertText

        protected override void DoInsertSpecialChar(IRtfInterpreterContext context, RtfVisualSpecialCharKind kind)
        {
            switch (kind)
            {
                case RtfVisualSpecialCharKind.Tabulator:
                    _plainText.Append(Settings.TabulatorText);
                    break;
                case RtfVisualSpecialCharKind.NonBreakingSpace:
                    _plainText.Append(Settings.NonBreakingSpaceText);
                    break;
                case RtfVisualSpecialCharKind.EmSpace:
                    _plainText.Append(Settings.EmSpaceText);
                    break;
                case RtfVisualSpecialCharKind.EnSpace:
                    _plainText.Append(Settings.EnSpaceText);
                    break;
                case RtfVisualSpecialCharKind.QmSpace:
                    _plainText.Append(Settings.QmSpaceText);
                    break;
                case RtfVisualSpecialCharKind.EmDash:
                    _plainText.Append(Settings.EmDashText);
                    break;
                case RtfVisualSpecialCharKind.EnDash:
                    _plainText.Append(Settings.EnDashText);
                    break;
                case RtfVisualSpecialCharKind.OptionalHyphen:
                    _plainText.Append(Settings.OptionalHyphenText);
                    break;
                case RtfVisualSpecialCharKind.NonBreakingHyphen:
                    _plainText.Append(Settings.NonBreakingHyphenText);
                    break;
                case RtfVisualSpecialCharKind.Bullet:
                    _plainText.Append(Settings.BulletText);
                    break;
                case RtfVisualSpecialCharKind.LeftSingleQuote:
                    _plainText.Append(Settings.LeftSingleQuoteText);
                    break;
                case RtfVisualSpecialCharKind.RightSingleQuote:
                    _plainText.Append(Settings.RightSingleQuoteText);
                    break;
                case RtfVisualSpecialCharKind.LeftDoubleQuote:
                    _plainText.Append(Settings.LeftDoubleQuoteText);
                    break;
                case RtfVisualSpecialCharKind.RightDoubleQuote:
                    _plainText.Append(Settings.RightDoubleQuoteText);
                    break;
                default:
                    _plainText.Append(Settings.UnknownSpecialCharText);
                    break;
            }
        } // DoInsertSpecialChar

        protected override void DoInsertBreak(IRtfInterpreterContext context, RtfVisualBreakKind kind)
        {
            switch (kind)
            {
                case RtfVisualBreakKind.Line:
                    _plainText.Append(Settings.LineBreakText);
                    break;
                case RtfVisualBreakKind.Page:
                    _plainText.Append(Settings.PageBreakText);
                    break;
                case RtfVisualBreakKind.Paragraph:
                    _plainText.Append(Settings.ParagraphBreakText);
                    break;
                case RtfVisualBreakKind.Section:
                    _plainText.Append(Settings.SectionBreakText);
                    break;
                default:
                    _plainText.Append(Settings.UnknownBreakText);
                    break;
            }
        } // DoInsertBreak

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

            _plainText.Append(imageText);
        } // DoInsertImage
    } // class RtfTextConverter
}