// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreter.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;

namespace Itenso.Rtf.Interpreter
{
    // ------------------------------------------------------------------------
    public sealed class RtfInterpreter : RtfInterpreterBase, IRtfElementVisitor
    {
        private readonly RtfColorTableBuilder colorTableBuilder;
        private readonly RtfDocumentInfoBuilder documentInfoBuilder;

        // ----------------------------------------------------------------------
        // members
        private readonly RtfFontTableBuilder fontTableBuilder;
        private readonly RtfImageBuilder imageBuilder;
        private readonly RtfUserPropertyBuilder userPropertyBuilder;
        private bool lastGroupWasPictureWrapper;

        // ----------------------------------------------------------------------
        public RtfInterpreter(params IRtfInterpreterListener[] listeners) :
            base(new RtfInterpreterSettings(), listeners)
        {
        } // RtfInterpreter

        // ----------------------------------------------------------------------
        public RtfInterpreter(IRtfInterpreterSettings settings, params IRtfInterpreterListener[] listeners) :
            base(settings, listeners)
        {
            fontTableBuilder = new RtfFontTableBuilder(Context.WritableFontTable, settings.IgnoreDuplicatedFonts);
            colorTableBuilder = new RtfColorTableBuilder(Context.WritableColorTable);
            documentInfoBuilder = new RtfDocumentInfoBuilder(Context.WritableDocumentInfo);
            userPropertyBuilder = new RtfUserPropertyBuilder(Context.WritableUserProperties);
            imageBuilder = new RtfImageBuilder();
        } // RtfInterpreter

        // ----------------------------------------------------------------------
        void IRtfElementVisitor.VisitTag(IRtfTag tag)
        {
            if (Context.State != RtfInterpreterState.InDocument)
                if (Context.FontTable.Count > 0)
                    if (Context.ColorTable.Count > 0 || RtfSpec.TagViewKind.Equals(tag.Name))
                        Context.State = RtfInterpreterState.InDocument;

            switch (Context.State)
            {
                case RtfInterpreterState.Init:
                    if (RtfSpec.TagRtf.Equals(tag.Name))
                    {
                        Context.State = RtfInterpreterState.InHeader;
                        Context.RtfVersion = tag.ValueAsNumber;
                    }
                    else
                    {
                        throw new RtfStructureException(Strings.InvalidInitTagState(tag.ToString()));
                    }
                    break;
                case RtfInterpreterState.InHeader:
                    switch (tag.Name)
                    {
                        case RtfSpec.TagDefaultFont:
                            Context.DefaultFontId = RtfSpec.TagFont + tag.ValueAsNumber;
                            break;
                    }
                    break;
                case RtfInterpreterState.InDocument:
                    switch (tag.Name)
                    {
                        case RtfSpec.TagPlain:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveNormal();
                            break;
                        case RtfSpec.TagParagraphDefaults:
                        case RtfSpec.TagSectionDefaults:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithAlignment(RtfTextAlignment.Left);
                            break;
                        case RtfSpec.TagBold:
                            var bold = !tag.HasValue || tag.ValueAsNumber != 0;
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithBold(bold);
                            break;
                        case RtfSpec.TagItalic:
                            var italic = !tag.HasValue || tag.ValueAsNumber != 0;
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithItalic(italic);
                            break;
                        case RtfSpec.TagUnderLine:
                            var underline = !tag.HasValue || tag.ValueAsNumber != 0;
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithUnderline(underline);
                            break;
                        case RtfSpec.TagUnderLineNone:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithUnderline(false);
                            break;
                        case RtfSpec.TagStrikeThrough:
                            var strikeThrough = !tag.HasValue || tag.ValueAsNumber != 0;
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithStrikeThrough(strikeThrough);
                            break;
                        case RtfSpec.TagHidden:
                            var hidden = !tag.HasValue || tag.ValueAsNumber != 0;
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithHidden(hidden);
                            break;
                        case RtfSpec.TagFont:
                            var fontId = tag.FullName;
                            if (Context.FontTable.ContainsFontWithId(fontId))
                            {
                                Context.WritableCurrentTextFormat =
                                    Context.WritableCurrentTextFormat.DeriveWithFont(
                                        Context.FontTable[fontId]);
                            }
                            else
                            {
                                if (Settings.IgnoreUnknownFonts && Context.FontTable.Count > 0)
                                    Context.WritableCurrentTextFormat =
                                        Context.WritableCurrentTextFormat.DeriveWithFont(Context.FontTable[0]);
                                else
                                    throw new RtfUndefinedFontException(Strings.UndefinedFont(fontId));
                            }
                            break;
                        case RtfSpec.TagFontSize:
                            var fontSize = tag.ValueAsNumber;
                            if (fontSize >= 0)
                                Context.WritableCurrentTextFormat =
                                    Context.WritableCurrentTextFormat.DeriveWithFontSize(fontSize);
                            else
                                throw new RtfInvalidDataException(Strings.InvalidFontSize(fontSize));
                            break;
                        case RtfSpec.TagFontSubscript:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithSuperScript(false);
                            break;
                        case RtfSpec.TagFontSuperscript:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithSuperScript(true);
                            break;
                        case RtfSpec.TagFontNoSuperSub:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithSuperScript(0);
                            break;
                        case RtfSpec.TagFontDown:
                            var moveDown = tag.ValueAsNumber;
                            if (moveDown == 0)
                                moveDown = 6; // the default value according to rtf spec
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithSuperScript(-moveDown);
                            break;
                        case RtfSpec.TagFontUp:
                            var moveUp = tag.ValueAsNumber;
                            if (moveUp == 0)
                                moveUp = 6; // the default value according to rtf spec
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithSuperScript(moveUp);
                            break;
                        case RtfSpec.TagAlignLeft:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithAlignment(RtfTextAlignment.Left);
                            break;
                        case RtfSpec.TagAlignCenter:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithAlignment(RtfTextAlignment.Center);
                            break;
                        case RtfSpec.TagAlignRight:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithAlignment(RtfTextAlignment.Right);
                            break;
                        case RtfSpec.TagAlignJustify:
                            Context.WritableCurrentTextFormat =
                                Context.WritableCurrentTextFormat.DeriveWithAlignment(RtfTextAlignment.Justify);
                            break;
                        case RtfSpec.TagColorBackground:
                        case RtfSpec.TagColorBackgroundWord:
                        case RtfSpec.TagColorHighlight:
                        case RtfSpec.TagColorForeground:
                            var colorIndex = tag.ValueAsNumber;
                            if (colorIndex >= 0 && colorIndex < Context.ColorTable.Count)
                            {
                                var newColor = Context.ColorTable[colorIndex];
                                var isForeground = RtfSpec.TagColorForeground.Equals(tag.Name);
                                Context.WritableCurrentTextFormat = isForeground
                                    ? Context.WritableCurrentTextFormat.DeriveWithForegroundColor(newColor)
                                    : Context.WritableCurrentTextFormat.DeriveWithBackgroundColor(newColor);
                            }
                            else
                            {
                                throw new RtfUndefinedColorException(Strings.UndefinedColor(colorIndex));
                            }
                            break;
                        case RtfSpec.TagSection:
                            NotifyInsertBreak(RtfVisualBreakKind.Section);
                            break;
                        case RtfSpec.TagParagraph:
                            NotifyInsertBreak(RtfVisualBreakKind.Paragraph);
                            break;
                        case RtfSpec.TagLine:
                            NotifyInsertBreak(RtfVisualBreakKind.Line);
                            break;
                        case RtfSpec.TagPage:
                            NotifyInsertBreak(RtfVisualBreakKind.Page);
                            break;
                        case RtfSpec.TagTabulator:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.Tabulator);
                            break;
                        case RtfSpec.TagTilde:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.NonBreakingSpace);
                            break;
                        case RtfSpec.TagEmDash:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.EmDash);
                            break;
                        case RtfSpec.TagEnDash:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.EnDash);
                            break;
                        case RtfSpec.TagEmSpace:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.EmSpace);
                            break;
                        case RtfSpec.TagEnSpace:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.EnSpace);
                            break;
                        case RtfSpec.TagQmSpace:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.QmSpace);
                            break;
                        case RtfSpec.TagBulltet:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.Bullet);
                            break;
                        case RtfSpec.TagLeftSingleQuote:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.LeftSingleQuote);
                            break;
                        case RtfSpec.TagRightSingleQuote:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.RightSingleQuote);
                            break;
                        case RtfSpec.TagLeftDoubleQuote:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.LeftDoubleQuote);
                            break;
                        case RtfSpec.TagRightDoubleQuote:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.RightDoubleQuote);
                            break;
                        case RtfSpec.TagHyphen:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.OptionalHyphen);
                            break;
                        case RtfSpec.TagUnderscore:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.NonBreakingHyphen);
                            break;
                    }
                    break;
            }
        } // IRtfElementVisitor.VisitTag

        // ----------------------------------------------------------------------
        void IRtfElementVisitor.VisitGroup(IRtfGroup group)
        {
            var groupDestination = group.Destination;
            switch (Context.State)
            {
                case RtfInterpreterState.Init:
                    if (RtfSpec.TagRtf.Equals(groupDestination))
                        VisitChildrenOf(group);
                    else
                        throw new RtfStructureException(Strings.InvalidInitGroupState(groupDestination));
                    break;
                case RtfInterpreterState.InHeader:
                    switch (groupDestination)
                    {
                        case RtfSpec.TagFontTable:
                            fontTableBuilder.VisitGroup(group);
                            break;
                        case RtfSpec.TagColorTable:
                            colorTableBuilder.VisitGroup(group);
                            break;
                        case RtfSpec.TagGenerator:
                            // last group with a destination in header, but no need to process its contents
                            Context.State = RtfInterpreterState.InDocument;
                            var generator = group.Contents.Count == 3 ? group.Contents[2] as IRtfText : null;
                            if (generator != null)
                            {
                                var generatorName = generator.Text;
                                Context.Generator = generatorName.EndsWith(";")
                                    ? generatorName.Substring(0, generatorName.Length - 1)
                                    : generatorName;
                            }
                            else
                            {
                                throw new RtfInvalidDataException(Strings.InvalidGeneratorGroup(group.ToString()));
                            }
                            break;
                        case RtfSpec.TagPlain:
                        case RtfSpec.TagParagraphDefaults:
                        case RtfSpec.TagSectionDefaults:
                        case RtfSpec.TagUnderLineNone:
                        case null:
                            // <tags>: special tags commonly used to reset state in a beginning group. necessary to recognize
                            //         state transition form header to document in case no other mechanism detects it and the
                            //         content starts with a group with such a 'destination' ...
                            // 'null': group without destination cannot be part of header, but need to process its contents
                            Context.State = RtfInterpreterState.InDocument;
                            if (!group.IsExtensionDestination)
                                VisitChildrenOf(group);
                            break;
                    }
                    break;
                case RtfInterpreterState.InDocument:
                    switch (groupDestination)
                    {
                        case RtfSpec.TagUserProperties:
                            userPropertyBuilder.VisitGroup(group);
                            break;
                        case RtfSpec.TagInfo:
                            documentInfoBuilder.VisitGroup(group);
                            break;
                        case RtfSpec.TagUnicodeAlternativeChoices:
                            var alternativeWithUnicodeSupport =
                                group.SelectChildGroupWithDestination(RtfSpec.TagUnicodeAlternativeUnicode);
                            if (alternativeWithUnicodeSupport != null)
                            {
                                // there is an alternative with unicode formatted content -> use this
                                VisitChildrenOf(alternativeWithUnicodeSupport);
                            }
                            else
                            {
                                // try to locate the alternative without unicode -> only ANSI fallbacks
                                var alternativeWithoutUnicode = // must be the third element if present
                                    group.Contents.Count > 2 ? group.Contents[2] as IRtfGroup : null;
                                if (alternativeWithoutUnicode != null)
                                    VisitChildrenOf(alternativeWithoutUnicode);
                            }
                            break;
                        case RtfSpec.TagHeader:
                        case RtfSpec.TagHeaderFirst:
                        case RtfSpec.TagHeaderLeft:
                        case RtfSpec.TagHeaderRight:
                        case RtfSpec.TagFooter:
                        case RtfSpec.TagFooterFirst:
                        case RtfSpec.TagFooterLeft:
                        case RtfSpec.TagFooterRight:
                        case RtfSpec.TagFootnote:
                        case RtfSpec.TagStyleSheet:
                            // groups we currently ignore, so their content doesn't intermix with
                            // the actual document content
                            break;
                        case RtfSpec.TagPictureWrapper:
                            VisitChildrenOf(group);
                            lastGroupWasPictureWrapper = true;
                            break;
                        case RtfSpec.TagPictureWrapperAlternative:
                            if (!lastGroupWasPictureWrapper)
                                VisitChildrenOf(group);
                            lastGroupWasPictureWrapper = false;
                            break;
                        case RtfSpec.TagPicture:
                            imageBuilder.VisitGroup(group);
                            NotifyInsertImage(
                                imageBuilder.Format,
                                imageBuilder.Width,
                                imageBuilder.Height,
                                imageBuilder.DesiredWidth,
                                imageBuilder.DesiredHeight,
                                imageBuilder.ScaleWidthPercent,
                                imageBuilder.ScaleHeightPercent,
                                imageBuilder.ImageDataHex);
                            break;
                        case RtfSpec.TagParagraphNumberText:
                        case RtfSpec.TagListNumberText:
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.ParagraphNumberBegin);
                            VisitChildrenOf(group);
                            NotifyInsertSpecialChar(RtfVisualSpecialCharKind.ParagraphNumberEnd);
                            break;
                        default:
                            if (!group.IsExtensionDestination)
                                VisitChildrenOf(group);
                            break;
                    }
                    break;
            }
        } // IRtfElementVisitor.VisitGroup

        // ----------------------------------------------------------------------
        void IRtfElementVisitor.VisitText(IRtfText text)
        {
            switch (Context.State)
            {
                case RtfInterpreterState.Init:
                    throw new RtfStructureException(Strings.InvalidInitTextState(text.Text));
                case RtfInterpreterState.InHeader:
                    // allow spaces in between header tables
                    if (!string.IsNullOrEmpty(text.Text.Trim()))
                        Context.State = RtfInterpreterState.InDocument;
                    break;
                case RtfInterpreterState.InDocument:
                    break;
            }
            NotifyInsertText(text.Text);
        } // IRtfElementVisitor.VisitText

        // ----------------------------------------------------------------------
        public static bool IsSupportedDocument(IRtfGroup rtfDocument)
        {
            try
            {
                GetSupportedDocument(rtfDocument);
            }
            catch (RtfException)
            {
                return false;
            }
            return true;
        } // IsSupportedDocument

        // ----------------------------------------------------------------------
        public static IRtfGroup GetSupportedDocument(IRtfGroup rtfDocument)
        {
            if (rtfDocument == null)
                throw new ArgumentNullException("rtfDocument");
            if (rtfDocument.Contents.Count == 0)
                throw new RtfEmptyDocumentException(Strings.EmptyDocument);
            var firstElement = rtfDocument.Contents[0];
            if (firstElement.Kind != RtfElementKind.Tag)
                throw new RtfStructureException(Strings.MissingDocumentStartTag);
            var firstTag = (IRtfTag) firstElement;
            if (!RtfSpec.TagRtf.Equals(firstTag.Name))
                throw new RtfStructureException(Strings.InvalidDocumentStartTag(RtfSpec.TagRtf));
            if (!firstTag.HasValue)
                throw new RtfUnsupportedStructureException(Strings.MissingRtfVersion);
            if (firstTag.ValueAsNumber != RtfSpec.RtfVersion1)
                throw new RtfUnsupportedStructureException(Strings.UnsupportedRtfVersion(firstTag.ValueAsNumber));
            return rtfDocument;
        } // GetSupportedDocument

        // ----------------------------------------------------------------------
        protected override void DoInterpret(IRtfGroup rtfDocument)
        {
            InterpretContents(GetSupportedDocument(rtfDocument));
        } // DoInterpret

        // ----------------------------------------------------------------------
        private void InterpretContents(IRtfGroup rtfDocument)
        {
            // by getting here we already know that the given document is supported, and hence
            // we know it has version 1
            Context.Reset(); // clears all previous content and sets the version to 1
            lastGroupWasPictureWrapper = false;
            NotifyBeginDocument();
            VisitChildrenOf(rtfDocument);
            Context.State = RtfInterpreterState.Ended;
            NotifyEndDocument();
        } // InterpretContents

        // ----------------------------------------------------------------------
        private void VisitChildrenOf(IRtfGroup group)
        {
            var pushedTextFormat = false;
            if (Context.State == RtfInterpreterState.InDocument)
            {
                Context.PushCurrentTextFormat();
                pushedTextFormat = true;
            }
            try
            {
                foreach (IRtfElement child in group.Contents)
                    child.Visit(this);
            }
            finally
            {
                if (pushedTextFormat)
                    Context.PopCurrentTextFormat();
            }
        } // VisitChildrenOf
    } // class RtfInterpreter
} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------