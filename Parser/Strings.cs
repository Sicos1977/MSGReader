// -- FILE ------------------------------------------------------------------
// name       : Strings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2009.02.19
// language   : c#
// environment: .NET 3.5
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Globalization;
using System.Resources;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf
{
    // ------------------------------------------------------------------------
    /// <summary>Provides strongly typed resource access for this namespace.</summary>
    internal sealed class Strings : StringsBase
    {
        // ----------------------------------------------------------------------
        // members
        private static readonly ResourceManager inst = NewInst(typeof(Strings));

        // ----------------------------------------------------------------------
        public static string ToManyBraces
        {
            get { return inst.GetString("ToManyBraces"); }
        } // ToManyBraces

        // ----------------------------------------------------------------------
        public static string ToFewBraces
        {
            get { return inst.GetString("ToFewBraces"); }
        } // ToFewBraces

        // ----------------------------------------------------------------------
        public static string NoRtfContent
        {
            get { return inst.GetString("NoRtfContent"); }
        } // NoRtfContent

        // ----------------------------------------------------------------------
        public static string UnexpectedEndOfFile
        {
            get { return inst.GetString("UnexpectedEndOfFile"); }
        } // UnexpectedEndOfFile

        // ----------------------------------------------------------------------
        public static string EndOfFileInvalidCharacter
        {
            get { return inst.GetString("EndOfFileInvalidCharacter"); }
        } // EndOfFileInvalidCharacter

        // ----------------------------------------------------------------------
        public static string MissingGroupForNewTag
        {
            get { return inst.GetString("MissingGroupForNewTag"); }
        } // MissingGroupForNewTag

        // ----------------------------------------------------------------------
        public static string MissingGroupForNewText
        {
            get { return inst.GetString("MissingGroupForNewText"); }
        } // MissingGroupForNewText

        // ----------------------------------------------------------------------
        public static string MultipleRootLevelGroups
        {
            get { return inst.GetString("MultipleRootLevelGroups"); }
        } // MultipleRootLevelGroups

        // ----------------------------------------------------------------------
        public static string UnclosedGroups
        {
            get { return inst.GetString("UnclosedGroups"); }
        } // UnclosedGroups

        // ----------------------------------------------------------------------
        public static string LogGroupBegin
        {
            get { return inst.GetString("LogGroupBegin"); }
        } // LogGroupBegin

        // ----------------------------------------------------------------------
        public static string LogGroupEnd
        {
            get { return inst.GetString("LogGroupEnd"); }
        } // LogGroupEnd

        // ----------------------------------------------------------------------
        public static string LogOverflowText
        {
            get { return inst.GetString("LogOverflowText"); }
        } // LogOverflowText

        // ----------------------------------------------------------------------
        public static string LogParseBegin
        {
            get { return inst.GetString("LogParseBegin"); }
        } // LogParseBegin

        // ----------------------------------------------------------------------
        public static string LogParseEnd
        {
            get { return inst.GetString("LogParseEnd"); }
        } // LogParseEnd

        // ----------------------------------------------------------------------
        public static string LogParseFail
        {
            get { return inst.GetString("LogParseFail"); }
        } // LogParseFail

        // ----------------------------------------------------------------------
        public static string LogParseFailUnknown
        {
            get { return inst.GetString("LogParseFailUnknown"); }
        } // LogParseFailUnknown

        // ----------------------------------------------------------------------
        public static string LogParseSuccess
        {
            get { return inst.GetString("LogParseSuccess"); }
        } // LogParseSuccess

        // ----------------------------------------------------------------------
        public static string LogTag
        {
            get { return inst.GetString("LogTag"); }
        } // LogTag

        // ----------------------------------------------------------------------
        public static string LogText
        {
            get { return inst.GetString("LogText"); }
        } // LogText

        // ----------------------------------------------------------------------
        public static string InvalidFirstHexDigit(char hexDigit)
        {
            return Format(inst.GetString("InvalidFirstHexDigit"), hexDigit);
        } // InvalidFirstHexDigit

        // ----------------------------------------------------------------------
        public static string InvalidSecondHexDigit(char hexDigit)
        {
            return Format(inst.GetString("InvalidSecondHexDigit"), hexDigit);
        } // InvalidSecondHexDigit

        // ----------------------------------------------------------------------
        public static string TagOnRootLevel(string tagName)
        {
            return Format(inst.GetString("TagOnRootLevel"), tagName);
        } // TagOnRootLevel

        // ----------------------------------------------------------------------
        public static string InvalidUnicodeSkipCount(string tagName)
        {
            return Format(inst.GetString("InvalidUnicodeSkipCount"), tagName);
        } // InvalidUnicodeSkipCount

        // ----------------------------------------------------------------------
        public static string InvalidMultiByteEncoding(byte[] buffer, int index, Encoding encoding)
        {
            var buf = new StringBuilder();
            for (var i = 0; i < index; i++)
                buf.Append(string.Format(
                    CultureInfo.InvariantCulture,
                    "{0:X}",
                    buffer[i]));
            return Format(inst.GetString("InvalidMultiByteEncoding"),
                buf.ToString(), encoding.EncodingName, encoding.CodePage);
        } // InvalidMultiByteEncoding

        // ----------------------------------------------------------------------
        public static string TextOnRootLevel(string text)
        {
            return Format(inst.GetString("TextOnRootLevel"), text);
        } // TextOnRootLevel
    } // class Strings
} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------