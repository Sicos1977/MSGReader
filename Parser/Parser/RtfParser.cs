// name       : RtfParser.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Collections;
using System.Globalization;
using System.IO;
using System.Text;
using Itenso.Rtf.Model;

namespace Itenso.Rtf.Parser
{
    public sealed class RtfParser : RtfParserBase
    {
        private readonly byte[] _byteDecodingBuffer = new byte[8]; // >0 for multi-byte encodings
        private readonly char[] _charDecodingBuffer = new char[1];
        private readonly Stack _codePageStack = new Stack();
        private readonly Hashtable _fontToCodePageMapping = new Hashtable();
        private readonly MemoryStream _hexDecodingBuffer = new MemoryStream();
        private readonly Stack _unicodeSkipCountStack = new Stack();
        private Decoder _byteToCharDecoder;

        // Members
        private StringBuilder _curText;
        private Encoding _encoding;
        private bool _expectingThemeFont;
        private int _fontTableStartLevel;
        private int _level;
        private int _tagCount;
        private int _tagCountAtLastGroupStart;
        private string _targetFont;
        private int _unicodeSkipCount;

        public RtfParser()
        {
        } // RtfParser

        public RtfParser(params IRtfParserListener[] listeners) :
            base(listeners)
        {
        } // RtfParser

        protected override void DoParse(IRtfSource rtfTextSource)
        {
            NotifyParseBegin();
            try
            {
                ParseRtf(rtfTextSource.Reader);
                NotifyParseSuccess();
            }
            catch (RtfException e)
            {
                NotifyParseFail(e);
                throw;
            }
            finally
            {
                NotifyParseEnd();
            }
        } // DoParse

        private void ParseRtf(TextReader reader)
        {
            _curText = new StringBuilder();

            _unicodeSkipCountStack.Clear();
            _codePageStack.Clear();
            _unicodeSkipCount = 1;
            _level = 0;
            _tagCountAtLastGroupStart = 0;
            _tagCount = 0;
            _fontTableStartLevel = -1;
            _targetFont = null;
            _expectingThemeFont = false;
            _fontToCodePageMapping.Clear();
            _hexDecodingBuffer.SetLength(0);
            UpdateEncoding(RtfSpec.AnsiCodePage);
            var groupCount = 0;
            const int eof = -1;
            var nextChar = PeekNextChar(reader, false);
            var backslashAlreadyConsumed = false;
            while (nextChar != eof)
            {
                var peekChar = 0;
                var peekCharValid = false;
                switch (nextChar)
                {
                    case '\\':
                        if (!backslashAlreadyConsumed)
                            reader.Read(); // must still consume the 'peek'ed char
                        var secondChar = PeekNextChar(reader, true);
                        switch (secondChar)
                        {
                            case '\\':
                            case '{':
                            case '}':
                                _curText.Append(ReadOneChar(reader)); // must still consume the 'peek'ed char
                                break;

                            case '\n':
                            case '\r':
                                reader.Read(); // must still consume the 'peek'ed char
                                // must be treated as a 'par' tag if preceded by a backslash
                                // (see RTF spec page 144)
                                HandleTag(reader, new RtfTag(RtfSpec.TagParagraph));
                                break;

                            case '\'':
                                reader.Read(); // must still consume the 'peek'ed char
                                var hex1 = (char) ReadOneByte(reader);
                                var hex2 = (char) ReadOneByte(reader);
                                if (!IsHexDigit(hex1))
                                    throw new RtfHexEncodingException(Strings.InvalidFirstHexDigit(hex1));
                                if (!IsHexDigit(hex2))
                                    throw new RtfHexEncodingException(Strings.InvalidSecondHexDigit(hex2));
                                var decodedByte = int.Parse("" + hex1 + hex2, NumberStyles.HexNumber);
                                _hexDecodingBuffer.WriteByte((byte) decodedByte);
                                peekChar = PeekNextChar(reader, false);
                                peekCharValid = true;
                                var mustFlushHexContent = true;
                                if (peekChar == '\\')
                                {
                                    reader.Read();
                                    backslashAlreadyConsumed = true;
                                    var continuationChar = PeekNextChar(reader, false);
                                    if (continuationChar == '\'')
                                        mustFlushHexContent = false;
                                }
                                if (mustFlushHexContent)
                                    DecodeCurrentHexBuffer();
                                break;

                            case '|':
                            case '~':
                            case '-':
                            case '_':
                            case ':':
                            case '*':
                                HandleTag(reader, new RtfTag("" + ReadOneChar(reader)));
                                // must still consume the 'peek'ed char
                                break;

                            default:
                                ParseTag(reader);
                                break;
                        }
                        break;

                    case '\n':
                    case '\r':
                        reader.Read(); // must still consume the 'peek'ed char
                        break;

                    case '\t':
                        reader.Read(); // must still consume the 'peek'ed char
                        // should be treated as a 'tab' tag (see RTF spec page 144)
                        HandleTag(reader, new RtfTag(RtfSpec.TagTabulator));
                        break;

                    case '{':
                        reader.Read(); // must still consume the 'peek'ed char
                        FlushText();
                        NotifyGroupBegin();
                        _tagCountAtLastGroupStart = _tagCount;
                        _unicodeSkipCountStack.Push(_unicodeSkipCount);
                        _codePageStack.Push(_encoding == null ? 0 : _encoding.CodePage);
                        _level++;
                        break;

                    case '}':
                        reader.Read(); // must still consume the 'peek'ed char
                        FlushText();
                        if (_level > 0)
                        {
                            _unicodeSkipCount = (int) _unicodeSkipCountStack.Pop();
                            if (_fontTableStartLevel == _level)
                            {
                                _fontTableStartLevel = -1;
                                _targetFont = null;
                                _expectingThemeFont = false;
                            }
                            UpdateEncoding((int) _codePageStack.Pop());
                            _level--;
                            NotifyGroupEnd();
                            groupCount++;
                        }
                        else
                        {
                            throw new RtfBraceNestingException(Strings.ToManyBraces);
                        }
                        break;

                    default:
                        _curText.Append(ReadOneChar(reader)); // must still consume the 'peek'ed char
                        break;
                }
                if (_level == 0 && IgnoreContentAfterRootGroup)
                    break;
                if (peekCharValid)
                {
                    nextChar = peekChar;
                }
                else
                {
                    nextChar = PeekNextChar(reader, false);
                    backslashAlreadyConsumed = false;
                }
            }
            FlushText();
            reader.Close();

            if (_level > 0)
                throw new RtfBraceNestingException(Strings.ToFewBraces);
            if (groupCount == 0)
                throw new RtfEmptyDocumentException(Strings.NoRtfContent);
            _curText = null;
        } // ParseRtf

        private void ParseTag(TextReader reader)
        {
            var tagName = new StringBuilder();
            StringBuilder tagValue = null;
            var readingName = true;
            var delimReached = false;

            var nextChar = PeekNextChar(reader, true);
            while (!delimReached)
            {
                if (readingName && IsAsciiLetter(nextChar))
                {
                    tagName.Append(ReadOneChar(reader)); // must still consume the 'peek'ed char
                }
                else if (IsDigit(nextChar) || nextChar == '-' && tagValue == null)
                {
                    readingName = false;
                    if (tagValue == null)
                        tagValue = new StringBuilder();
                    tagValue.Append(ReadOneChar(reader)); // must still consume the 'peek'ed char
                }
                else
                {
                    delimReached = true;
                    IRtfTag newTag;
                    if (tagValue != null && tagValue.Length > 0)
                        newTag = new RtfTag(tagName.ToString(), tagValue.ToString());
                    else
                        newTag = new RtfTag(tagName.ToString());
                    var skippedContent = HandleTag(reader, newTag);
                    if (nextChar == ' ' && !skippedContent)
                        reader.Read(); // must still consume the 'peek'ed char
                }
                if (!delimReached)
                    nextChar = PeekNextChar(reader, true);
            }
        } // ParseTag

        private bool HandleTag(TextReader reader, IRtfTag tag)
        {
            if (_level == 0)
                throw new RtfStructureException(Strings.TagOnRootLevel(tag.ToString()));

            if (_tagCount < 4)
                UpdateEncoding(tag);

            var tagName = tag.Name;
            // enable the font name detection in case the last tag was introducing
            // a theme font
            var detectFontName = _expectingThemeFont;
            if (_tagCountAtLastGroupStart == _tagCount)
            {
                // first tag in a group
                switch (tagName)
                {
                    case RtfSpec.TagThemeFontLoMajor:
                    case RtfSpec.TagThemeFontHiMajor:
                    case RtfSpec.TagThemeFontDbMajor:
                    case RtfSpec.TagThemeFontBiMajor:
                    case RtfSpec.TagThemeFontLoMinor:
                    case RtfSpec.TagThemeFontHiMinor:
                    case RtfSpec.TagThemeFontDbMinor:
                    case RtfSpec.TagThemeFontBiMinor:
                        // these introduce a new font, but the actual font tag will be
                        // the second tag in the group, so we must remember this condition ...
                        _expectingThemeFont = true;
                        break;
                }
                // always enable the font name detection also for the first tag in a group
                detectFontName = true;
            }
            if (detectFontName)
                switch (tagName)
                {
                    case RtfSpec.TagFont:
                        if (_fontTableStartLevel > 0)
                        {
                            // in the font-table definition:
                            _targetFont = tag.FullName;
                            _expectingThemeFont = false; // reset that state now
                        }
                        break;
                    case RtfSpec.TagFontTable:
                        _fontTableStartLevel = _level;
                        break;
                }
            if (_targetFont != null)
                if (RtfSpec.TagFontCharset.Equals(tagName))
                {
                    var charSet = tag.ValueAsNumber;
                    var codePage = RtfSpec.GetCodePage(charSet);
                    _fontToCodePageMapping[_targetFont] = codePage;
                    UpdateEncoding(codePage);
                }
            if (_fontToCodePageMapping.Count > 0 && RtfSpec.TagFont.Equals(tagName))
            {
                var codePage = (int?) _fontToCodePageMapping[tag.FullName];
                if (codePage != null)
                    UpdateEncoding(codePage.Value);
            }

            var skippedContent = false;
            switch (tagName)
            {
                case RtfSpec.TagUnicodeCode:
                    var unicodeValue = tag.ValueAsNumber;
                    var unicodeChar = (char) unicodeValue;
                    _curText.Append(unicodeChar);
                    // skip over the indicated number of 'alternative representation' text
                    for (var i = 0; i < _unicodeSkipCount; i++)
                    {
                        var nextChar = PeekNextChar(reader, true);
                        switch (nextChar)
                        {
                            case ' ':
                            case '\r':
                            case '\n':
                                reader.Read(); // consume peeked char
                                skippedContent = true;
                                if (i == 0)
                                    i--;
                                break;
                            case '\\':
                                reader.Read(); // consume peeked char
                                skippedContent = true;
                                var secondChar = ReadOneByte(reader); // mandatory
                                switch (secondChar)
                                {
                                    case '\'':
                                        // ok, this is a hex-encoded 'byte' -> need to consume both
                                        // hex digits too
                                        ReadOneByte(reader); // high nibble
                                        ReadOneByte(reader); // low nibble
                                        break;
                                }
                                break;
                            case '{':
                            case '}':
                                // don't consume peeked char and abort skipping
                                i = _unicodeSkipCount;
                                break;
                            default:
                                reader.Read(); // consume peeked char
                                skippedContent = true;
                                break;
                        }
                    }
                    break;

                case RtfSpec.TagUnicodeSkipCount:
                    var newSkipCount = tag.ValueAsNumber;
                    if (newSkipCount < 0 || newSkipCount > 10)
                        throw new RtfUnicodeEncodingException(Strings.InvalidUnicodeSkipCount(tag.ToString()));
                    _unicodeSkipCount = newSkipCount;
                    break;

                default:
                    FlushText();
                    NotifyTagFound(tag);
                    break;
            }

            _tagCount++;

            return skippedContent;
        } // HandleTag

        private void UpdateEncoding(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagEncodingAnsi:
                    UpdateEncoding(RtfSpec.AnsiCodePage);
                    break;
                case RtfSpec.TagEncodingMac:
                    UpdateEncoding(10000);
                    break;
                case RtfSpec.TagEncodingPc:
                    UpdateEncoding(437);
                    break;
                case RtfSpec.TagEncodingPca:
                    UpdateEncoding(850);
                    break;
                case RtfSpec.TagEncodingAnsiCodePage:
                    UpdateEncoding(tag.ValueAsNumber);
                    break;
            }
        } // UpdateEncoding

        private void UpdateEncoding(int codePage)
        {
            if (_encoding == null || codePage != _encoding.CodePage)
            {
                switch (codePage)
                {
                    case RtfSpec.AnsiCodePage:
                    case RtfSpec.SymbolFakeCodePage: // hack to handle a windows legacy ...
                        _encoding = RtfSpec.AnsiEncoding;
                        break;
                    default:
                        _encoding = Encoding.GetEncoding(codePage);
                        break;
                }
                _byteToCharDecoder = null;
            }
            if (_byteToCharDecoder == null)
                _byteToCharDecoder = _encoding.GetDecoder();
        } // UpdateEncoding

        private static bool IsAsciiLetter(int character)
        {
            return character >= 'a' && character <= 'z' || character >= 'A' && character <= 'Z';
        } // IsASCIILetter

        private static bool IsHexDigit(int character)
        {
            return character >= '0' && character <= '9' ||
                   character >= 'a' && character <= 'f' ||
                   character >= 'A' && character <= 'F';
        } // IsHexDigit

        private static bool IsDigit(int character)
        {
            return character >= '0' && character <= '9';
        } // IsDigit

        private static int ReadOneByte(TextReader reader)
        {
            var byteValue = reader.Read();
            if (byteValue == -1)
                throw new RtfUnicodeEncodingException(Strings.UnexpectedEndOfFile);
            return byteValue;
        } // ReadOneByte

        private char ReadOneChar(TextReader reader)
        {
            // NOTE: the handling of multi-byte encodings is probably not the most
            // efficient here ...

            var completed = false;
            var byteIndex = 0;
            while (!completed)
            {
                _byteDecodingBuffer[byteIndex] = (byte) ReadOneByte(reader);
                byteIndex++;
                int usedBytes;
                int usedChars;
                _byteToCharDecoder.Convert(
                    _byteDecodingBuffer, 0, byteIndex,
                    _charDecodingBuffer, 0, 1,
                    true,
                    out usedBytes,
                    out usedChars,
                    out completed);
                if (completed && (usedBytes != byteIndex || usedChars != 1))
                    throw new RtfMultiByteEncodingException(Strings.InvalidMultiByteEncoding(
                        _byteDecodingBuffer, byteIndex, _encoding));
            }
            var character = _charDecodingBuffer[0];
            return character;
        } // ReadOneChar

        private void DecodeCurrentHexBuffer()
        {
            var pendingByteCount = _hexDecodingBuffer.Length;
            if (pendingByteCount > 0)
            {
                var pendingBytes = _hexDecodingBuffer.ToArray();
                var convertedChars = new char[pendingByteCount]; // should be enough

                var startIndex = 0;
                var completed = false;
                while (!completed && startIndex < pendingBytes.Length)
                {
                    int usedBytes;
                    int usedChars;
                    _byteToCharDecoder.Convert(
                        pendingBytes, startIndex, pendingBytes.Length - startIndex,
                        convertedChars, 0, convertedChars.Length,
                        true,
                        out usedBytes,
                        out usedChars,
                        out completed);
                    _curText.Append(convertedChars, 0, usedChars);
                    startIndex += usedChars;
                }

                _hexDecodingBuffer.SetLength(0);
            }
        } // DecodeCurrentHexBuffer

        // ReSharper disable UnusedParameter.Local
        private static int PeekNextChar(TextReader reader, bool mandatory)
// ReSharper restore UnusedParameter.Local
        {
            var character = reader.Peek();
            if (mandatory && character == -1)
                throw new RtfMultiByteEncodingException(Strings.EndOfFileInvalidCharacter);
            return character;
        } // PeekNextChar

        private void FlushText()
        {
            if (_curText.Length > 0)
            {
                NotifyTextFound(new RtfText(_curText.ToString()));
                _curText.Remove(0, _curText.Length);
            }
        } // FlushText
    } // class RtfParser
}