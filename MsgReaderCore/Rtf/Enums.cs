namespace MsgReader.Rtf
{
    #region Enum RtfAlignment
    /// <summary>
    /// Text alignment
    /// </summary>
    internal enum RtfAlignment
    {
        /// <summary>
        /// left
        /// </summary>
        Left,
        
        /// <summary>
        /// center
        /// </summary>
        Center,
        
        /// <summary>
        /// right
        /// </summary>
        Right,
        
        /// <summary>
        /// justify
        /// </summary>
        Justify
    }
    #endregion

    #region Enum RtfPictureType
    internal enum RtfPictureType
    {
        /// <summary>
        /// Source of the picture is an EMF (enhanced metafile).
        /// </summary>
        Emfblip,
        
        /// <summary>
        /// Source of the picture is a PNG.
        /// </summary>
        Pngblip,
        
        /// <summary>
        /// Source of the picture is a JPEG.
        /// </summary>
        Jpegblip,
        
        /// <summary>
        /// ource of the picture is QuickDraw.
        /// </summary>
        Macpict,
        
        /// <summary>
        /// Source of the picture is an OS/2 metafile. The N argument identifies the metafile type. The N values are described in the \pmmetafile table further on in this section.
        /// </summary>
        Pmmetafile,
        
        /// <summary>
        /// Source of the picture is a Windows metafile. The N argument identifies the metafile type (the default type is 1).
        /// </summary>
        Wmetafile,
        
        /// <summary>
        /// Source of the picture is a Windows device-independent bitmap. The N argument identifies the bitmap type, which must equal 0.
        /// The information to be included in RTF from a Windows device-independent bitmap is the concatenation of the BITMAPINFO structure followed by the actual pixel data.
        /// </summary>
        Dibitmap,
        
        /// <summary>
        /// Source of the picture is a Windows device-dependent bitmap. The N argument identifies the bitmap type (must equal 0).
        /// The information to be included in RTF from a Windows device-dependent bitmap is the result of the GetBitmapBits function.
        /// </summary>
        Wbitmap
    }
    #endregion

    #region Enum RtfObjectType
    internal enum RtfObjectType
    {
        Emb,
        Link,
        AutLink,
        Sub,
        Pub,
        Icemb,
        Html,
        Ocx
    }
    #endregion

    #region Enum RtfTokenType
    /// <summary>
    /// Rtf token type
    /// </summary>
    internal enum RtfTokenType
    {
        None,
        Keyword,
        ExtKeyword,
        Control,
        Text,
        Eof,
        GroupStart,
        GroupEnd
    }
    #endregion

    #region Enum RtfDomFieldMethod
    /// <summary>
    /// Rtf dom field method
    /// </summary>
    internal enum RtfDomFieldMethod
    {
        None,
        Dirty,
        Edit,
        Lock,
        Priv,
    }
    #endregion

    #region Enum RtfNodeType
    internal enum RtfNodeType
    {
        /// <summary>
        /// root
        /// </summary>
        Root,
        /// <summary>
        /// keyword, etc /marginl
        /// </summary>
        Keyword,
        /// <summary>
        /// external keyword node , etc. /*/keyword
        /// </summary>
        ExtKeyword,
        /// <summary>
        /// control
        /// </summary>
        Control,
        /// <summary>
        /// plain text
        /// </summary>
        Text,
        /// <summary>
        /// group , etc . { }
        /// </summary>
        Group,
        /// <summary>
        /// nothing
        /// </summary>
        None
    }
    #endregion

    #region Enum RtfVerticalAlignment
    /// <summary>
    /// Rtf vertical alignment
    /// </summary>
    internal enum RtfVerticalAlignment
    {
        /// <summary>
        /// top alignment
        /// </summary>
        Top,
        /// <summary>
        /// middle alignment
        /// </summary>
        Middle,
        /// <summary>
        /// bottom alignment
        /// </summary>
        Bottom
    }
    #endregion

    #region Enum RtfLevelNumberType
    internal enum RtfLevelNumberType
    {
        None = -10,

        ///<summary>Arabic (1, 2, 3)</summary>
        Arabic = 0,

        ///<summary>Uppercase Roman numeral (I, II, III)</summary>
        UppercaseRomanNumeral = 1,

        ///<summary>Lowercase Roman numeral (i, ii, iii)</summary>
        LowercaseRomanNumeral = 2,

        ///<summary>Uppercase letter (A, B, C)</summary>
        UppercaseLetter = 3,

        ///<summary>Lowercase letter (a, b, c)</summary>
        LowercaseLetter = 4,

        ///<summary>Ordinal number (1st, 2nd, 3rd)</summary>
        OrdinalNumber = 5,

        ///<summary>Cardinal text number (One, Two Three)</summary>
        CardinalTextNumber = 6,

        ///<summary>Ordinal text number (First, Second, Third)</summary>
        OrdinalTextNumber = 7,

        ///<summary>Kanji numbering without the digit character (*dbnum1)</summary>
        KanjiNumberingWithoutTheDigitCharacter = 10,

        ///<summary>Kanji numbering with the digit character (*dbnum2)</summary>
        KanjiNumberingWithTheDigitCharacte = 11,

        ///<summary>46 phonetic double_byte katakana characters (*aiueo*dbchar)</summary>
        // ReSharper disable InconsistentNaming
        _46_phonetic_double_byte_katakana_characters_aiueo_dbchar = 20,

        ///<summary>46 phonetic double_byte katakana characters (*iroha*dbchar)</summary>
        _46_phonetic_double_byte_katakana_characters_iroha_dbchar = 21,
        ///<summary>46 phonetic katakana characters in "aiueo" order (*aiueo)</summary>
        // ReSharper disable once InconsistentNaming
        _46_phonetic_katakana_characters_in_aiueo_order = 12,

        ///<summary>46 phonetic katakana characters in "iroha" order (*iroha)</summary>
        // ReSharper disable once InconsistentNaming
        _46_phonetic_katakana_characters_in_iroha_order = 13,
        // ReSharper restore InconsistentNaming

        ///<summary>Double_byte character</summary>
        DoubleByteCharacter = 14,

        ///<summary>Single_byte character</summary>
        SingleByteCharacter = 15,

        ///<summary>Kanji numbering 3 (*dbnum3)</summary>
        KanjiNumbering3 = 16,

        ///<summary>Kanji numbering 4 (*dbnum4)</summary>
        KanjiNumbering4 = 17,

        ///<summary>Circle numbering (*circlenum)</summary>
        CircleNumbering = 18,

        ///<summary>Double_byte Arabic numbering</summary>
        DoubleByteArabicNumbering = 19,

        ///<summary>Arabic with leading zero (01, 02, 03, ..., 10, 11)</summary>
        ArabicWithLeadingZero = 22,

        ///<summary>Bullet (no number at all)</summary>
        Bullet = 23,

        ///<summary>Korean numbering 2 (*ganada)</summary>
        KoreanNumbering2 = 24,

        ///<summary>Korean numbering 1 (*chosung)</summary>
        KoreanNumbering1 = 25,

        ///<summary>Chinese numbering 1 (*gb1)</summary>
        ChineseNumbering1 = 26,

        ///<summary>Chinese numbering 2 (*gb2)</summary>
        ChineseNumbering2 = 27,

        ///<summary>Chinese numbering 3 (*gb3)</summary>
        ChineseNumbering3 = 28,

        ///<summary>Chinese numbering 4 (*gb4)</summary>
        ChineseNumbering4 = 29,

        ///<summary>Chinese Zodiac numbering 1 (* zodiac1)</summary>
        ChineseZodiacNumbering1 = 30,

        ///<summary>Chinese Zodiac numbering 2 (* zodiac2) </summary>
        ChineseZodiacNumbering2 = 31,

        ///<summary>Chinese Zodiac numbering 3 (* zodiac3)</summary>
        ChineseZodiacNumbering3 = 32,

        ///<summary>Taiwanese double_byte numbering 1</summary>
        TaiwaneseDoubleByteNumbering1 = 33,

        ///<summary>Taiwanese double_byte numbering 2</summary>
        TaiwaneseDoubleByteNumbering2 = 34,

        ///<summary>Taiwanese double_byte numbering 3</summary>
        TaiwaneseDoubleByteNumbering3 = 35,

        ///<summary>Taiwanese double_byte numbering 4</summary>
        TaiwaneseDoubleByteNumbering4 = 36,

        ///<summary>Chinese double_byte numbering 1</summary>
        ChineseDoubleByteNumbering1 = 37,

        ///<summary>Chinese double_byte numbering 2</summary>
        ChineseDoubleByteNumbering2 = 38,

        ///<summary>Chinese double_byte numbering 3</summary>
        ChineseDoubleByteNumbering3 = 39,

        ///<summary>Chinese double_byte numbering 4</summary>
        ChineseDoubleByteNumbering4 = 40,

        ///<summary>Korean double_byte numbering 1</summary>
        KoreanDoubleByteNumbering1 = 41,

        ///<summary>Korean double_byte numbering 2</summary>
        KoreanDoubleByteNumbering2 = 42,

        ///<summary>Korean double_byte numbering 3</summary>
        KoreanDoubleByteNumbering3 = 43,

        ///<summary>Korean double_byte numbering 4</summary>
        KoreanDoubleByteNumbering4 = 44,

        ///<summary>Hebrew non_standard decimal </summary>
        HebrewNonStandardDecimal = 45,

        ///<summary>Arabic Alif Ba Tah</summary>
        ArabicAlifBaTah = 46,

        ///<summary>Hebrew Biblical standard</summary>
        HebrewBiblicalStandard = 47,

        ///<summary>Arabic Abjad style</summary>
        ArabicAbjadStyle = 48,

        ///<summary>No number</summary>
        NoNumber = 255
    }
    #endregion

    #region Enum RtfHeaderFooterStyle
    internal enum RtfHeaderFooterStyle
    {
        AllPages,
        LeftPages,
        RightPages,
        FirstPage
    }
    #endregion
}
