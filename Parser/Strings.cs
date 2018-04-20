// name       : Strings.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2009.02.19
// language   : c#
// environment: .NET 3.5
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Globalization;
using System.Resources;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf
{
    /// <summary>Provides strongly typed resource access for this namespace.</summary>
    internal sealed class Strings : StringsBase
    {
        // Members
        private static readonly ResourceManager Inst = NewInst(typeof(Strings));

        public static string ToManyBraces
        {
            get { return Inst.GetString("ToManyBraces"); }
        } // ToManyBraces

        public static string ToFewBraces
        {
            get { return Inst.GetString("ToFewBraces"); }
        } // ToFewBraces

        public static string NoRtfContent
        {
            get { return Inst.GetString("NoRtfContent"); }
        } // NoRtfContent

        public static string UnexpectedEndOfFile
        {
            get { return Inst.GetString("UnexpectedEndOfFile"); }
        } // UnexpectedEndOfFile

        public static string EndOfFileInvalidCharacter
        {
            get { return Inst.GetString("EndOfFileInvalidCharacter"); }
        } // EndOfFileInvalidCharacter

        public static string MissingGroupForNewTag
        {
            get { return Inst.GetString("MissingGroupForNewTag"); }
        } // MissingGroupForNewTag

        public static string MissingGroupForNewText
        {
            get { return Inst.GetString("MissingGroupForNewText"); }
        } // MissingGroupForNewText

        public static string MultipleRootLevelGroups
        {
            get { return Inst.GetString("MultipleRootLevelGroups"); }
        } // MultipleRootLevelGroups

        public static string UnclosedGroups
        {
            get { return Inst.GetString("UnclosedGroups"); }
        } // UnclosedGroups

        public static string LogGroupBegin
        {
            get { return Inst.GetString("LogGroupBegin"); }
        } // LogGroupBegin

        public static string LogGroupEnd
        {
            get { return Inst.GetString("LogGroupEnd"); }
        } // LogGroupEnd

        public static string LogOverflowText
        {
            get { return Inst.GetString("LogOverflowText"); }
        } // LogOverflowText

        public static string LogParseBegin
        {
            get { return Inst.GetString("LogParseBegin"); }
        } // LogParseBegin

        public static string LogParseEnd
        {
            get { return Inst.GetString("LogParseEnd"); }
        } // LogParseEnd

        public static string LogParseFail
        {
            get { return Inst.GetString("LogParseFail"); }
        } // LogParseFail

        public static string LogParseFailUnknown
        {
            get { return Inst.GetString("LogParseFailUnknown"); }
        } // LogParseFailUnknown

        public static string LogParseSuccess
        {
            get { return Inst.GetString("LogParseSuccess"); }
        } // LogParseSuccess

        public static string LogTag
        {
            get { return Inst.GetString("LogTag"); }
        } // LogTag

        public static string LogText
        {
            get { return Inst.GetString("LogText"); }
        } // LogText

        public static string InvalidFirstHexDigit(char hexDigit)
        {
            return Format(Inst.GetString("InvalidFirstHexDigit"), hexDigit);
        } // InvalidFirstHexDigit

        public static string InvalidSecondHexDigit(char hexDigit)
        {
            return Format(Inst.GetString("InvalidSecondHexDigit"), hexDigit);
        } // InvalidSecondHexDigit

        public static string TagOnRootLevel(string tagName)
        {
            return Format(Inst.GetString("TagOnRootLevel"), tagName);
        } // TagOnRootLevel

        public static string InvalidUnicodeSkipCount(string tagName)
        {
            return Format(Inst.GetString("InvalidUnicodeSkipCount"), tagName);
        } // InvalidUnicodeSkipCount

        public static string InvalidMultiByteEncoding(byte[] buffer, int index, Encoding encoding)
        {
            var buf = new StringBuilder();
            for (var i = 0; i < index; i++)
                buf.Append(string.Format(
                    CultureInfo.InvariantCulture,
                    "{0:X}",
                    buffer[i]));
            return Format(Inst.GetString("InvalidMultiByteEncoding"),
                buf.ToString(), encoding.EncodingName, encoding.CodePage);
        } // InvalidMultiByteEncoding

        public static string TextOnRootLevel(string text)
        {
            return Format(Inst.GetString("TextOnRootLevel"), text);
        } // TextOnRootLevel
    } // class Strings
}