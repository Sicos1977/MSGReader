// name       : Strings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2009.02.19
// language   : c#
// environment: .NET 3.5
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Resources;
using Itenso.Sys;

namespace Itenso.Rtf
{
    /// <summary>Provides strongly typed resource access for this namespace.</summary>
    internal sealed class Strings : StringsBase
    {
        // Members
        private static readonly ResourceManager Inst = NewInst(typeof(Strings));

        public static string EmptyDocument
        {
            get { return Inst.GetString("EmptyDocument"); }
        } // EmptyDocument

        public static string MissingDocumentStartTag
        {
            get { return Inst.GetString("MissingDocumentStartTag"); }
        } // MissingDocumentStartTag

        public static string MissingRtfVersion
        {
            get { return Inst.GetString("MissingRtfVersion"); }
        } // MissingRtfVersion

        public static string InvalidTextContextState
        {
            get { return Inst.GetString("InvalidTextContextState"); }
        } // InvalidTextContextState

        public static string ImageFormatText
        {
            get { return Inst.GetString("ImageFormatText"); }
        } // ImageFormatText

        public static string LogBeginDocument
        {
            get { return Inst.GetString("LogBeginDocument"); }
        } // LogBeginDocument

        public static string LogEndDocument
        {
            get { return Inst.GetString("LogEndDocument"); }
        } // LogEndDocument

        public static string LogInsertBreak
        {
            get { return Inst.GetString("LogInsertBreak"); }
        } // LogInsertBreak

        public static string LogInsertChar
        {
            get { return Inst.GetString("LogInsertChar"); }
        } // LogInsertChar

        public static string LogInsertImage
        {
            get { return Inst.GetString("LogInsertImage"); }
        } // LogInsertImage

        public static string LogInsertText
        {
            get { return Inst.GetString("LogInsertText"); }
        } // LogInsertText

        public static string LogOverflowText
        {
            get { return Inst.GetString("LogOverflowText"); }
        } // LogOverflowText

        public static string ColorTableUnsupportedText(string text)
        {
            return Format(Inst.GetString("ColorTableUnsupportedText"), text);
        } // ColorTableUnsupportedText

        public static string DuplicateFont(string fontId)
        {
            return Format(Inst.GetString("DuplicateFont"), fontId);
        } // DuplicateFont

        public static string InvalidDocumentStartTag(string expectedTag)
        {
            return Format(Inst.GetString("InvalidDocumentStartTag"), expectedTag);
        } // InvalidDocumentStartTag

        public static string InvalidInitTagState(string tag)
        {
            return Format(Inst.GetString("InvalidInitTagState"), tag);
        } // InvalidInitTagState

        public static string UndefinedFont(string fontId)
        {
            return Format(Inst.GetString("UndefinedFont"), fontId);
        } // UndefinedFont

        public static string InvalidFontSize(int fontSize)
        {
            return Format(Inst.GetString("InvalidFontSize"), fontSize);
        } // InvalidFontSize

        public static string UndefinedColor(int colorIndex)
        {
            return Format(Inst.GetString("UndefinedColor"), colorIndex);
        } // UndefinedColor

        public static string InvalidInitGroupState(string group)
        {
            return Format(Inst.GetString("InvalidInitGroupState"), group);
        } // InvalidInitGroupState

        public static string InvalidGeneratorGroup(string group)
        {
            return Format(Inst.GetString("InvalidGeneratorGroup"), group);
        } // InvalidGeneratorGroup

        public static string InvalidInitTextState(string text)
        {
            return Format(Inst.GetString("InvalidInitTextState"), text);
        } // InvalidInitTextState	

        public static string InvalidDefaultFont(string fontId, string allowedFontIds)
        {
            return Format(Inst.GetString("InvalidDefaultFont"), fontId, allowedFontIds);
        } // InvalidDefaultFont

        public static string UnsupportedRtfVersion(int version)
        {
            return Format(Inst.GetString("UnsupportedRtfVersion"), version);
        } // UnsupportedRtfVersion

        public static string InvalidColor(int color)
        {
            return Format(Inst.GetString("InvalidColor"), color);
        } // InvalidColor

        public static string InvalidCharacterSet(int charSet)
        {
            return Format(Inst.GetString("InvalidCharacterSet"), charSet);
        } // InvalidCharacterSet

        public static string InvalidCodePage(int codePage)
        {
            return Format(Inst.GetString("InvalidCodePage"), codePage);
        } // InvalidCodePage

        public static string FontSizeOutOfRange(int fontSize)
        {
            return Format(Inst.GetString("FontSizeOutOfRange"), fontSize);
        } // FontSizeOutOfRange

        public static string InvalidImageWidth(int width)
        {
            return Format(Inst.GetString("InvalidImageWidth"), width);
        } // InvalidImageWidth

        public static string InvalidImageHeight(int height)
        {
            return Format(Inst.GetString("InvalidImageHeight"), height);
        } // InvalidImageHeight

        public static string InvalidImageDesiredHeight(int width)
        {
            return Format(Inst.GetString("InvalidImageDesiredHeight"), width);
        } // InvalidImageDesiredHeight

        public static string InvalidImageDesiredWidth(int height)
        {
            return Format(Inst.GetString("InvalidImageDesiredWidth"), height);
        } // InvalidImageDesiredWidth

        public static string InvalidImageScaleWidth(int scaleWidth)
        {
            return Format(Inst.GetString("InvalidImageScaleWidth"), scaleWidth);
        } // InvalidImageScaleWidth

        public static string InvalidImageScaleHeight(int scaleHeight)
        {
            return Format(Inst.GetString("InvalidImageScaleHeight"), scaleHeight);
        } // InvalidImageScaleHeight
    } // class Strings
}