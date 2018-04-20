// -- FILE ------------------------------------------------------------------
// name       : RtfSpec.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Text;

namespace Itenso.Rtf
{
    // ------------------------------------------------------------------------
    public static class RtfSpec
    {
        // --- rtf general ----
        public const string TagRtf = "rtf";
        public const int RtfVersion1 = 1;

        public const string TagGenerator = "generator";
        public const string TagViewKind = "viewkind";

        // --- encoding ----
        public const string TagEncodingAnsi = "ansi";
        public const string TagEncodingMac = "mac";
        public const string TagEncodingPc = "pc";
        public const string TagEncodingPca = "pca";
        public const string TagEncodingAnsiCodePage = "ansicpg";
        public const int AnsiCodePage = 1252;
        public const int SymbolFakeCodePage = 42; // a windows legacy hack ...

        public const string TagUnicodeSkipCount = "uc";
        public const string TagUnicodeCode = "u";
        public const string TagUnicodeAlternativeChoices = "upr";
        public const string TagUnicodeAlternativeUnicode = "ud";

        // --- font ----
        public const string TagFontTable = "fonttbl";
        public const string TagDefaultFont = "deff";
        public const string TagFont = "f";
        public const string TagFontKindNil = "fnil";
        public const string TagFontKindRoman = "froman";
        public const string TagFontKindSwiss = "fswiss";
        public const string TagFontKindModern = "fmodern";
        public const string TagFontKindScript = "fscript";
        public const string TagFontKindDecor = "fdecor";
        public const string TagFontKindTech = "ftech";
        public const string TagFontKindBidi = "fbidi";
        public const string TagFontCharset = "fcharset";
        public const string TagFontPitch = "fprq";
        public const string TagFontSize = "fs";
        public const string TagFontDown = "dn";
        public const string TagFontUp = "up";
        public const string TagFontSubscript = "sub";
        public const string TagFontSuperscript = "super";
        public const string TagFontNoSuperSub = "nosupersub";

        public const string TagThemeFontLoMajor = "flomajor"; // these are 'theme' fonts
        public const string TagThemeFontHiMajor = "fhimajor"; // used in new font tables
        public const string TagThemeFontDbMajor = "fdbmajor";
        public const string TagThemeFontBiMajor = "fbimajor";
        public const string TagThemeFontLoMinor = "flominor";
        public const string TagThemeFontHiMinor = "fhiminor";
        public const string TagThemeFontDbMinor = "fdbminor";
        public const string TagThemeFontBiMinor = "fbiminor";

        public const int DefaultFontSize = 24;

        public const string TagCodePage = "cpg";

        // --- color ----
        public const string TagColorTable = "colortbl";
        public const string TagColorRed = "red";
        public const string TagColorGreen = "green";
        public const string TagColorBlue = "blue";
        public const string TagColorForeground = "cf";
        public const string TagColorBackground = "cb";
        public const string TagColorBackgroundWord = "chcbpat";
        public const string TagColorHighlight = "highlight";

        // --- header/footer ----
        public const string TagHeader = "header";
        public const string TagHeaderFirst = "headerf";
        public const string TagHeaderLeft = "headerl";
        public const string TagHeaderRight = "headerr";
        public const string TagFooter = "footer";
        public const string TagFooterFirst = "footerf";
        public const string TagFooterLeft = "footerl";
        public const string TagFooterRight = "footerr";
        public const string TagFootnote = "footnote";

        // --- character ----
        public const string TagDelimiter = ";";
        public const string TagExtensionDestination = "*";
        public const string TagTilde = "~";
        public const string TagHyphen = "-";
        public const string TagUnderscore = "_";

        // --- special character ----
        public const string TagPage = "page";
        public const string TagSection = "sect";
        public const string TagParagraph = "par";
        public const string TagLine = "line";
        public const string TagTabulator = "tab";
        public const string TagEmDash = "emdash";
        public const string TagEnDash = "endash";
        public const string TagEmSpace = "emspace";
        public const string TagEnSpace = "enspace";
        public const string TagQmSpace = "qmspace";
        public const string TagBulltet = "bullet";
        public const string TagLeftSingleQuote = "lquote";
        public const string TagRightSingleQuote = "rquote";
        public const string TagLeftDoubleQuote = "ldblquote";
        public const string TagRightDoubleQuote = "rdblquote";

        // --- format ----
        public const string TagPlain = "plain";
        public const string TagParagraphDefaults = "pard";
        public const string TagSectionDefaults = "sectd";

        public const string TagBold = "b";
        public const string TagItalic = "i";
        public const string TagUnderLine = "ul";
        public const string TagUnderLineNone = "ulnone";
        public const string TagStrikeThrough = "strike";
        public const string TagHidden = "v";
        public const string TagAlignLeft = "ql";
        public const string TagAlignCenter = "qc";
        public const string TagAlignRight = "qr";
        public const string TagAlignJustify = "qj";

        public const string TagStyleSheet = "stylesheet";

        // --- info ----
        public const string TagInfo = "info";
        public const string TagInfoVersion = "version";
        public const string TagInfoRevision = "vern";
        public const string TagInfoNumberOfPages = "nofpages";
        public const string TagInfoNumberOfWords = "nofwords";
        public const string TagInfoNumberOfChars = "nofchars";
        public const string TagInfoId = "id";
        public const string TagInfoTitle = "title";
        public const string TagInfoSubject = "subject";
        public const string TagInfoAuthor = "author";
        public const string TagInfoManager = "manager";
        public const string TagInfoCompany = "company";
        public const string TagInfoOperator = "operator";
        public const string TagInfoCategory = "category";
        public const string TagInfoKeywords = "keywords";
        public const string TagInfoComment = "comment";
        public const string TagInfoDocumentComment = "doccomm";
        public const string TagInfoHyperLinkBase = "hlinkbase";
        public const string TagInfoCreationTime = "creatim";
        public const string TagInfoRevisionTime = "revtim";
        public const string TagInfoPrintTime = "printim";
        public const string TagInfoBackupTime = "buptim";
        public const string TagInfoYear = "yr";
        public const string TagInfoMonth = "mo";
        public const string TagInfoDay = "dy";
        public const string TagInfoHour = "hr";
        public const string TagInfoMinute = "min";
        public const string TagInfoSecond = "sec";
        public const string TagInfoEditingTimeMinutes = "edmins";

        // --- user properties ----
        public const string TagUserProperties = "userprops";
        public const string TagUserPropertyType = "proptype";
        public const string TagUserPropertyName = "propname";
        public const string TagUserPropertyValue = "staticval";
        public const string TagUserPropertyLink = "linkval";

        // this table is from the RTF specification 1.9.1, page 40
        public const int PropertyTypeInteger = 3;
        public const int PropertyTypeRealNumber = 5;
        public const int PropertyTypeDate = 64;
        public const int PropertyTypeBoolean = 11;
        public const int PropertyTypeText = 30;

        // --- picture ----
        public const string TagPicture = "pict";
        public const string TagPictureWrapper = "shppict";
        public const string TagPictureWrapperAlternative = "nonshppict";
        public const string TagPictureFormatEmf = "emfblip";
        public const string TagPictureFormatPng = "pngblip";
        public const string TagPictureFormatJpg = "jpegblip";
        public const string TagPictureFormatPict = "macpict";
        public const string TagPictureFormatOs2Metafile = "pmmetafile";
        public const string TagPictureFormatWmf = "wmetafile";
        public const string TagPictureFormatWinDib = "dibitmap";
        public const string TagPictureFormatWinBmp = "wbitmap";
        public const string TagPictureWidth = "picw";
        public const string TagPictureHeight = "pich";
        public const string TagPictureWidthGoal = "picwgoal";
        public const string TagPictureHeightGoal = "pichgoal";
        public const string TagPictureWidthScale = "picscalex";
        public const string TagPictureHeightScale = "picscaley";

        // --- bullets/numbering ----
        public const string TagParagraphNumberText = "pntext";
        public const string TagListNumberText = "listtext";
        public static readonly Encoding AnsiEncoding = Encoding.GetEncoding(AnsiCodePage);

        // ----------------------------------------------------------------------
        public static int GetCodePage(int charSet)
        {
            switch (charSet)
            {
                case 0:
                    return 1252; // ANSI
                case 1:
                    return 0; // Default
                case 2:
                    return 42; // Symbol
                case 77:
                    return 10000; // Mac Roman
                case 78:
                    return 10001; // Mac Shift Jis
                case 79:
                    return 10003; // Mac Hangul
                case 80:
                    return 10008; // Mac GB2312
                case 81:
                    return 10002; // Mac Big5
                case 82:
                    return 0; // Mac Johab (old)
                case 83:
                    return 10005; // Mac Hebrew
                case 84:
                    return 10004; // Mac Arabic
                case 85:
                    return 10006; // Mac Greek
                case 86:
                    return 10081; // Mac Turkish
                case 87:
                    return 10021; // Mac Thai
                case 88:
                    return 10029; // Mac East Europe
                case 89:
                    return 10007; // Mac Russian
                case 128:
                    return 932; // Shift JIS
                case 129:
                    return 949; // Hangul
                case 130:
                    return 1361; // Johab
                case 134:
                    return 936; // GB2312
                case 136:
                    return 950; // Big5
                case 161:
                    return 1253; // Greek
                case 162:
                    return 1254; // Turkish
                case 163:
                    return 1258; // Vietnamese
                case 177:
                    return 1255; // Hebrew
                case 178:
                    return 1256; // Arabic
                case 179:
                    return 0; // Arabic Traditional (old)
                case 180:
                    return 0; // Arabic user (old)
                case 181:
                    return 0; // Hebrew user (old)
                case 186:
                    return 1257; // Baltic
                case 204:
                    return 1251; // Russian
                case 222:
                    return 874; // Thai
                case 238:
                    return 1250; // Eastern European
                case 254:
                    return 437; // PC 437
                case 255:
                    return 850; // OEM
            }

            return 0;
        } // GetCodePage
    } // class RtfSpec
} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------