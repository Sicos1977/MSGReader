// name       : RtfInterpreterListenerDocumentBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Text;
using Itenso.Rtf.Model;

namespace Itenso.Rtf.Interpreter
{
    public sealed class RtfInterpreterListenerDocumentBuilder : RtfInterpreterListenerBase
    {
        private readonly RtfVisualCollection _pendingParagraphContent = new RtfVisualCollection();
        private readonly StringBuilder _pendingText = new StringBuilder();

        // Members

        private RtfDocument _document;

        private IRtfTextFormat _pendingTextFormat;
        private RtfVisualCollection _visualDocumentContent;

        public bool CombineTextWithSameFormat { get; set; } = true;

// CombineTextWithSameFormat

        public IRtfDocument Document
        {
            get { return _document; }
        } // Document

        protected override void DoBeginDocument(IRtfInterpreterContext context)
        {
            _document = null;
            _visualDocumentContent = new RtfVisualCollection();
        } // DoBeginDocument

        protected override void DoInsertText(IRtfInterpreterContext context, string text)
        {
            if (CombineTextWithSameFormat)
            {
                var newFormat = context.GetSafeCurrentTextFormat();
                if (!newFormat.Equals(_pendingTextFormat))
                    FlushPendingText();
                _pendingTextFormat = newFormat;
                _pendingText.Append(text);
            }
            else
            {
                AppendAlignedVisual(new RtfVisualText(text, context.GetSafeCurrentTextFormat()));
            }
        } // DoInsertText

        protected override void DoInsertSpecialChar(IRtfInterpreterContext context, RtfVisualSpecialCharKind kind)
        {
            FlushPendingText();
            _visualDocumentContent.Add(new RtfVisualSpecialChar(kind));
        } // DoInsertSpecialChar

        protected override void DoInsertBreak(IRtfInterpreterContext context, RtfVisualBreakKind kind)
        {
            FlushPendingText();
            _visualDocumentContent.Add(new RtfVisualBreak(kind));
            switch (kind)
            {
                case RtfVisualBreakKind.Paragraph:
                case RtfVisualBreakKind.Section:
                    EndParagraph(context);
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
            FlushPendingText();
            AppendAlignedVisual(new RtfVisualImage(format,
                context.GetSafeCurrentTextFormat().Alignment,
                width, height, desiredWidth, desiredHeight,
                scaleWidthPercent, scaleHeightPercent, imageDataHex));
        } // DoInsertImage

        protected override void DoEndDocument(IRtfInterpreterContext context)
        {
            FlushPendingText();
            EndParagraph(context);
            _document = new RtfDocument(context, _visualDocumentContent);
            _visualDocumentContent = null;
            _visualDocumentContent = null;
        } // DoEndDocument

        private void EndParagraph(IRtfInterpreterContext context)
        {
            var finalParagraphAlignment = context.GetSafeCurrentTextFormat().Alignment;
            foreach (IRtfVisual alignedVisual in _pendingParagraphContent)
                switch (alignedVisual.Kind)
                {
                    case RtfVisualKind.Image:
                        var image = (RtfVisualImage) alignedVisual;
                        // ReSharper disable RedundantCheckBeforeAssignment
                        if (image.Alignment != finalParagraphAlignment)
                            // ReSharper restore RedundantCheckBeforeAssignment
                            image.Alignment = finalParagraphAlignment;
                        break;
                    case RtfVisualKind.Text:
                        var text = (RtfVisualText) alignedVisual;
                        if (text.Format.Alignment != finalParagraphAlignment)
                        {
                            IRtfTextFormat correctedFormat =
                                ((RtfTextFormat) text.Format).DeriveWithAlignment(finalParagraphAlignment);
                            var correctedUniqueFormat = context.GetUniqueTextFormatInstance(correctedFormat);
                            text.Format = correctedUniqueFormat;
                        }
                        break;
                }
            _pendingParagraphContent.Clear();
        } // EndParagraph

        private void FlushPendingText()
        {
            if (_pendingTextFormat != null)
            {
                AppendAlignedVisual(new RtfVisualText(_pendingText.ToString(), _pendingTextFormat));
                _pendingTextFormat = null;
                _pendingText.Remove(0, _pendingText.Length);
            }
        } // FlushPendingText

        private void AppendAlignedVisual(RtfVisual visual)
        {
            _visualDocumentContent.Add(visual);
            _pendingParagraphContent.Add(visual);
        } // AppendAlignedVisual
    } // class RtfInterpreterListenerDocumentBuilder
}