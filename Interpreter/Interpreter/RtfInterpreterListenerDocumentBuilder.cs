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
        private readonly RtfVisualCollection pendingParagraphContent = new RtfVisualCollection();
        private readonly StringBuilder pendingText = new StringBuilder();

        // members

        private RtfDocument document;

        private IRtfTextFormat pendingTextFormat;
        private RtfVisualCollection visualDocumentContent;

        public bool CombineTextWithSameFormat { get; set; } = true;

// CombineTextWithSameFormat

        public IRtfDocument Document
        {
            get { return document; }
        } // Document

        protected override void DoBeginDocument(IRtfInterpreterContext context)
        {
            document = null;
            visualDocumentContent = new RtfVisualCollection();
        } // DoBeginDocument

        protected override void DoInsertText(IRtfInterpreterContext context, string text)
        {
            if (CombineTextWithSameFormat)
            {
                var newFormat = context.GetSafeCurrentTextFormat();
                if (!newFormat.Equals(pendingTextFormat))
                    FlushPendingText();
                pendingTextFormat = newFormat;
                pendingText.Append(text);
            }
            else
            {
                AppendAlignedVisual(new RtfVisualText(text, context.GetSafeCurrentTextFormat()));
            }
        } // DoInsertText

        protected override void DoInsertSpecialChar(IRtfInterpreterContext context, RtfVisualSpecialCharKind kind)
        {
            FlushPendingText();
            visualDocumentContent.Add(new RtfVisualSpecialChar(kind));
        } // DoInsertSpecialChar

        protected override void DoInsertBreak(IRtfInterpreterContext context, RtfVisualBreakKind kind)
        {
            FlushPendingText();
            visualDocumentContent.Add(new RtfVisualBreak(kind));
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
            document = new RtfDocument(context, visualDocumentContent);
            visualDocumentContent = null;
            visualDocumentContent = null;
        } // DoEndDocument

        private void EndParagraph(IRtfInterpreterContext context)
        {
            var finalParagraphAlignment = context.GetSafeCurrentTextFormat().Alignment;
            foreach (IRtfVisual alignedVisual in pendingParagraphContent)
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
            pendingParagraphContent.Clear();
        } // EndParagraph

        private void FlushPendingText()
        {
            if (pendingTextFormat != null)
            {
                AppendAlignedVisual(new RtfVisualText(pendingText.ToString(), pendingTextFormat));
                pendingTextFormat = null;
                pendingText.Remove(0, pendingText.Length);
            }
        } // FlushPendingText

        private void AppendAlignedVisual(RtfVisual visual)
        {
            visualDocumentContent.Add(visual);
            pendingParagraphContent.Add(visual);
        } // AppendAlignedVisual
    } // class RtfInterpreterListenerDocumentBuilder
}