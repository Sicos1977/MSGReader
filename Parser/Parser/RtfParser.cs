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
        private readonly byte[] byteDecodingBuffer = new byte[8]; // >0 for multi-byte encodings
        private readonly char[] charDecodingBuffer = new char[1];
        private readonly Stack codePageStack = new Stack();
        private readonly Hashtable fontToCodePageMapping = new Hashtable();
        private readonly MemoryStream hexDecodingBuffer = new MemoryStream();
        private readonly Stack unicodeSkipCountStack = new Stack();
        private Decoder byteToCharDecoder;

        // members
        private StringBuilder curText;
        private Encoding encoding;
        private bool expectingThemeFont;
        private int fontTableStartLevel;
        private int level;
        private int tagCount;
        private int tagCountAtLastGroupStart;
        private string targetFont;
        private int unicodeSkipCount;

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
            curText = new StringBuilder();

            unicodeSkipCountStack.Clear();
            codePageStack.Clear();
            unicodeSkipCount = 1;
            level = 0;
            tagCountAtLastGroupStart = 0;
            tagCount = 0;
            fontTableStartLevel = -1;
            targetFont = null;
            expectingThemeFont = false;
            fontToCodePageMapping.Clear();
            hexDecodingBuffer.SetLength(0);
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
                                curText.Append(ReadOneChar(reader)); // must still consume the 'peek'ed char
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
                                hexDecodingBuffer.WriteByte((byte) decodedByte);
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
                        tagCountAtLastGroupStart = tagCount;
                        unicodeSkipCountStack.Push(unicodeSkipCount);
                        codePageStack.Push(encoding == null ? 0 : encoding.CodePage);
                        level++;
                        break;

                    case '}':
                        reader.Read(); // must still consume the 'peek'ed char
                        FlushText();
                        if (level > 0)
                        {
                            unicodeSkipCount = (int) unicodeSkipCountStack.Pop();
                            if (fontTableStartLevel == level)
                            {
                                fontTableStartLevel = -1;
                                targetFont = null;
                                expectingThemeFont = false;
                            }
                            UpdateEncoding((int) codePageStack.Pop());
                            level--;
                            NotifyGroupEnd();
                            groupCount++;
                        }
                        else
                        {
                            throw new RtfBraceNestingException(Strings.ToManyBraces);
                        }
                        break;

                    default:
                        curText.Append(ReadOneChar(reader)); // must still consume the 'peek'ed char
                        break;
                }
                if (level == 0 && IgnoreContentAfterRootGroup)
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

            if (level > 0)
                throw new RtfBraceNestingException(Strings.ToFewBraces);
            if (groupCount == 0)
                throw new RtfEmptyDocumentException(Strings.NoRtfContent);
            curText = null;
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
                if (readingName && IsASCIILetter(nextChar))
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
            if (level == 0)
                throw new RtfStructureException(Strings.TagOnRootLevel(tag.ToString()));

            if (tagCount < 4)
                UpdateEncoding(tag);

            var tagName = tag.Name;
            // enable the font name detection in case the last tag was introducing
            // a theme font
            var detectFontName = expectingThemeFont;
            if (tagCountAtLastGroupStart == tagCount)
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
                        expectingThemeFont = true;
                        break;
                }
                // always enable the font name detection also for the first tag in a group
                detectFontName = true;
            }
            if (detectFontName)
                switch (tagName)
                {
                    case RtfSpec.TagFont:
                        if (fontTableStartLevel > 0)
                        {
                            // in the font-table definition:
                            targetFont = tag.FullName;
                            expectingThemeFont = false; // reset that state now
                        }
                        break;
                    case RtfSpec.TagFontTable:
                        fontTableStartLevel = level;
                        break;
                }
            if (targetFont != null)
                if (RtfSpec.TagFontCharset.Equals(tagName))
                {
                    var charSet = tag.ValueAsNumber;
                    var codePage = RtfSpec.GetCodePage(charSet);
                    fontToCodePageMapping[targetFont] = codePage;
                    UpdateEncoding(codePage);
                }
            if (fontToCodePageMapping.Count > 0 && RtfSpec.TagFont.Equals(tagName))
            {
                var codePage = (int?) fontToCodePageMapping[tag.FullName];
                if (codePage != null)
                    UpdateEncoding(codePage.Value);
            }

            var skippedContent = false;
            switch (tagName)
            {
                case RtfSpec.TagUnicodeCode:
                    var unicodeValue = tag.ValueAsNumber;
                    var unicodeChar = (char) unicodeValue;
                    curText.Append(unicodeChar);
                    // skip over the indicated number of 'alternative representation' text
                    for (var i = 0; i < unicodeSkipCount; i++)
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
                                i = unicodeSkipCount;
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
                    unicodeSkipCount = newSkipCount;
                    break;

                default:
                    FlushText();
                    NotifyTagFound(tag);
                    break;
            }

            tagCount++;

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
            if (encoding == null || codePage != encoding.CodePage)
            {
                switch (codePage)
                {
                    case RtfSpec.AnsiCodePage:
                    case RtfSpec.SymbolFakeCodePage: // hack to handle a windows legacy ...
                        encoding = RtfSpec.AnsiEncoding;
                        break;
                    default:
                        encoding = Encoding.GetEncoding(codePage);
                        break;
                }
                byteToCharDecoder = null;
            }
            if (byteToCharDecoder == null)
                byteToCharDecoder = encoding.GetDecoder();
        } // UpdateEncoding

        private static bool IsASCIILetter(int character)
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
                byteDecodingBuffer[byteIndex] = (byte) ReadOneByte(reader);
                byteIndex++;
                int usedBytes;
                int usedChars;
                byteToCharDecoder.Convert(
                    byteDecodingBuffer, 0, byteIndex,
                    charDecodingBuffer, 0, 1,
                    true,
                    out usedBytes,
                    out usedChars,
                    out completed);
                if (completed && (usedBytes != byteIndex || usedChars != 1))
                    throw new RtfMultiByteEncodingException(Strings.InvalidMultiByteEncoding(
                        byteDecodingBuffer, byteIndex, encoding));
            }
            var character = charDecodingBuffer[0];
            return character;
        } // ReadOneChar

        private void DecodeCurrentHexBuffer()
        {
            var pendingByteCount = hexDecodingBuffer.Length;
            if (pendingByteCount > 0)
            {
                var pendingBytes = hexDecodingBuffer.ToArray();
                var convertedChars = new char[pendingByteCount]; // should be enough

                var startIndex = 0;
                var completed = false;
                while (!completed && startIndex < pendingBytes.Length)
                {
                    int usedBytes;
                    int usedChars;
                    byteToCharDecoder.Convert(
                        pendingBytes, startIndex, pendingBytes.Length - startIndex,
                        convertedChars, 0, convertedChars.Length,
                        true,
                        out usedBytes,
                        out usedChars,
                        out completed);
                    curText.Append(convertedChars, 0, usedChars);
                    startIndex += usedChars;
                }

                hexDecodingBuffer.SetLength(0);
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
            if (curText.Length > 0)
            {
                NotifyTextFound(new RtfText(curText.ToString()));
                curText.Remove(0, curText.Length);
            }
        } // FlushText
    } // class RtfParser
}