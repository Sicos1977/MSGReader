// -- FILE ------------------------------------------------------------------
// name       : RtfParser.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Text;
using Itenso.Rtf.Model;

namespace Itenso.Rtf.Parser
{

	// ------------------------------------------------------------------------
	public sealed class RtfParser : RtfParserBase
	{

		// ----------------------------------------------------------------------
		public RtfParser()
		{
		} // RtfParser

		// ----------------------------------------------------------------------
		public RtfParser( params IRtfParserListener[] listeners ) :
			base( listeners )
		{
		} // RtfParser

		// ----------------------------------------------------------------------
		protected override void DoParse( IRtfSource rtfTextSource )
		{
			NotifyParseBegin();
			try
			{
				ParseRtf( rtfTextSource.Reader );
				NotifyParseSuccess();
			}
			catch ( RtfException e )
			{
				NotifyParseFail( e );
				throw;
			}
			finally
			{
				NotifyParseEnd();
			}
		} // DoParse

		// ----------------------------------------------------------------------
		private void ParseRtf( TextReader reader )
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
			hexDecodingBuffer.SetLength( 0 );
			UpdateEncoding( RtfSpec.AnsiCodePage );
			int groupCount = 0;
			const int eof = -1;
			int nextChar = PeekNextChar( reader, false );
			bool backslashAlreadyConsumed = false;
			while ( nextChar != eof )
			{
				int peekChar = 0;
				bool peekCharValid = false;
				switch ( nextChar )
				{
					case '\\':
						if ( !backslashAlreadyConsumed )
						{
							reader.Read(); // must still consume the 'peek'ed char
						}
						int secondChar = PeekNextChar( reader, true );
						switch ( secondChar )
						{
							case '\\':
							case '{':
							case '}':
								curText.Append( ReadOneChar( reader ) ); // must still consume the 'peek'ed char
								break;

							case '\n':
							case '\r':
								reader.Read(); // must still consume the 'peek'ed char
								// must be treated as a 'par' tag if preceded by a backslash
								// (see RTF spec page 144)
								HandleTag( reader, new RtfTag( RtfSpec.TagParagraph ) );
								break;

							case '\'':
								reader.Read(); // must still consume the 'peek'ed char
								char hex1 = (char)ReadOneByte( reader );
								char hex2 = (char)ReadOneByte( reader );
								if ( !IsHexDigit( hex1 ) )
								{
									throw new RtfHexEncodingException( Strings.InvalidFirstHexDigit( hex1 ) );
								}
								if ( !IsHexDigit( hex2 ) )
								{
									throw new RtfHexEncodingException( Strings.InvalidSecondHexDigit( hex2 ) );
								}
								int decodedByte = Int32.Parse( "" + hex1 + hex2, NumberStyles.HexNumber );
								hexDecodingBuffer.WriteByte( (byte)decodedByte );
								peekChar = PeekNextChar( reader, false );
								peekCharValid = true;
								bool mustFlushHexContent = true;
								if ( peekChar == '\\' )
								{
									reader.Read();
									backslashAlreadyConsumed = true;
									int continuationChar = PeekNextChar( reader, false );
									if ( continuationChar == '\'' )
									{
										mustFlushHexContent = false;
									}
								}
								if ( mustFlushHexContent )
								{
									// we may _NOT_ handle hex content in a character-by-character way as
									// this results in invalid text for japanese/chinese content ...
									// -> we wait until the following content is non-hex and then flush the
									//    pending data. ugly but necessary with our decoding model.
									DecodeCurrentHexBuffer();
								}
								break;

							case '|':
							case '~':
							case '-':
							case '_':
							case ':':
							case '*':
								HandleTag( reader, new RtfTag( "" + ReadOneChar( reader ) ) ); // must still consume the 'peek'ed char
								break;

							default:
								ParseTag( reader );
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
						HandleTag( reader, new RtfTag( RtfSpec.TagTabulator ) );
						break;

					case '{':
						reader.Read(); // must still consume the 'peek'ed char
						FlushText();
						NotifyGroupBegin();
						tagCountAtLastGroupStart = tagCount;
						unicodeSkipCountStack.Push( unicodeSkipCount );
						codePageStack.Push( encoding == null ? 0 : encoding.CodePage );
						level++;
						break;

					case '}':
						reader.Read(); // must still consume the 'peek'ed char
						FlushText();
						if ( level > 0 )
						{
							unicodeSkipCount = (int)unicodeSkipCountStack.Pop();
							if ( fontTableStartLevel == level )
							{
								fontTableStartLevel = -1;
								targetFont = null;
								expectingThemeFont = false;
							}
							UpdateEncoding( (int)codePageStack.Pop() );
							level--;
							NotifyGroupEnd();
							groupCount++;
						}
						else
						{
							throw new RtfBraceNestingException( Strings.ToManyBraces );
						}
						break;

					default:
						curText.Append( ReadOneChar( reader ) ); // must still consume the 'peek'ed char
						break;
				}
				if ( level == 0 && IgnoreContentAfterRootGroup )
				{
					break;
				}
				if ( peekCharValid )
				{
					nextChar = peekChar;
				}
				else
				{
					nextChar = PeekNextChar( reader, false );
					backslashAlreadyConsumed = false;
				}
			}
			FlushText();
			reader.Close();

			if ( level > 0 )
			{
				throw new RtfBraceNestingException( Strings.ToFewBraces );
			}
			if ( groupCount == 0 )
			{
				throw new RtfEmptyDocumentException( Strings.NoRtfContent );
			}
			curText = null;
		} // ParseRtf

		// ----------------------------------------------------------------------
		private void ParseTag( TextReader reader )
		{
			StringBuilder tagName = new StringBuilder();
			StringBuilder tagValue = null;
			bool readingName = true;
			bool delimReached = false;

			int nextChar = PeekNextChar( reader, true );
			while ( !delimReached )
			{
				if ( readingName && IsASCIILetter( nextChar ) )
				{
					tagName.Append( ReadOneChar( reader ) ); // must still consume the 'peek'ed char
				}
				else if ( IsDigit( nextChar ) || (nextChar == '-' && tagValue == null) )
				{
					readingName = false;
					if ( tagValue == null )
					{
						tagValue = new StringBuilder();
					}
					tagValue.Append( ReadOneChar( reader ) ); // must still consume the 'peek'ed char
				}
				else
				{
					delimReached = true;
					IRtfTag newTag;
					if ( tagValue != null && tagValue.Length > 0 )
					{
						newTag = new RtfTag( tagName.ToString(), tagValue.ToString() );
					}
					else
					{
						newTag = new RtfTag( tagName.ToString() );
					}
					bool skippedContent = HandleTag( reader, newTag );
					if ( nextChar == ' ' && !skippedContent )
					{
						reader.Read(); // must still consume the 'peek'ed char
					}
				}
				if ( !delimReached )
				{
					nextChar = PeekNextChar( reader, true );
				}
			}
		} // ParseTag

		// ----------------------------------------------------------------------
		private bool HandleTag( TextReader reader, IRtfTag tag )
		{
			if ( level == 0 )
			{
				throw new RtfStructureException( Strings.TagOnRootLevel( tag.ToString() ) );
			}

			if ( tagCount < 4 )
			{
				// this only handles the initial encoding tag in the header section
				UpdateEncoding( tag );
			}

			string tagName = tag.Name;
			// enable the font name detection in case the last tag was introducing
			// a theme font
			bool detectFontName = expectingThemeFont;
			if ( tagCountAtLastGroupStart == tagCount )
			{
				// first tag in a group
				switch ( tagName )
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
			if ( detectFontName )
			{
				// first tag in a group:
				switch ( tagName )
				{
					case RtfSpec.TagFont:
						if ( fontTableStartLevel > 0 )
						{
							// in the font-table definition:
							// -> remember the target font for charset mapping
							targetFont = tag.FullName;
							expectingThemeFont = false; // reset that state now
						}
						break;
					case RtfSpec.TagFontTable:
						// -> remember we're in the font-table definition
						fontTableStartLevel = level;
						break;
				}
			}
			if ( targetFont != null )
			{
				// within a font-tables font definition: perform charset mapping
				if ( RtfSpec.TagFontCharset.Equals( tagName ) )
				{
					int charSet = tag.ValueAsNumber;
					int codePage = RtfSpec.GetCodePage( charSet );
					fontToCodePageMapping[ targetFont ] = codePage;
					UpdateEncoding( codePage );
				}
			}
			if ( fontToCodePageMapping.Count > 0 && RtfSpec.TagFont.Equals( tagName ) )
			{
				int? codePage = (int?)fontToCodePageMapping[ tag.FullName ];
				if ( codePage != null )
				{
					UpdateEncoding( codePage.Value );
				}
			}

			bool skippedContent = false;
			switch ( tagName )
			{
				case RtfSpec.TagUnicodeCode:
					int unicodeValue = tag.ValueAsNumber;
					char unicodeChar = (char)unicodeValue;
					curText.Append( unicodeChar );
					// skip over the indicated number of 'alternative representation' text
					for ( int i = 0; i < unicodeSkipCount; i++ )
					{
						int nextChar = PeekNextChar( reader, true );
						switch ( nextChar )
						{
							case ' ':
							case '\r':
							case '\n':
								reader.Read(); // consume peeked char
								skippedContent = true;
								if ( i == 0 )
								{
									// the first whitespace after the tag
									// -> only a delimiter, doesn't count for skipping ...
									i--;
								}
								break;
							case '\\':
								reader.Read(); // consume peeked char
								skippedContent = true;
								int secondChar = ReadOneByte( reader ); // mandatory
								switch ( secondChar )
								{
									case '\'':
										// ok, this is a hex-encoded 'byte' -> need to consume both
										// hex digits too
										ReadOneByte( reader ); // high nibble
										ReadOneByte( reader ); // low nibble
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
					int newSkipCount = tag.ValueAsNumber;
					if ( newSkipCount < 0 || newSkipCount > 10 )
					{
						throw new RtfUnicodeEncodingException( Strings.InvalidUnicodeSkipCount( tag.ToString() ) );
					}
					unicodeSkipCount = newSkipCount;
					break;

				default:
					FlushText();
					NotifyTagFound( tag );
					break;
			}

			tagCount++;

			return skippedContent;
		} // HandleTag

		// ----------------------------------------------------------------------
		private void UpdateEncoding( IRtfTag tag )
		{
			switch ( tag.Name )
			{
				case RtfSpec.TagEncodingAnsi:
					UpdateEncoding( RtfSpec.AnsiCodePage );
					break;
				case RtfSpec.TagEncodingMac:
					UpdateEncoding( 10000 );
					break;
				case RtfSpec.TagEncodingPc:
					UpdateEncoding( 437 );
					break;
				case RtfSpec.TagEncodingPca:
					UpdateEncoding( 850 );
					break;
				case RtfSpec.TagEncodingAnsiCodePage:
					UpdateEncoding( tag.ValueAsNumber );
					break;
			}
		} // UpdateEncoding

		// ----------------------------------------------------------------------
		private void UpdateEncoding( int codePage )
		{
			if ( encoding == null || codePage != encoding.CodePage )
			{
				switch ( codePage )
				{
					case RtfSpec.AnsiCodePage:
					case RtfSpec.SymbolFakeCodePage: // hack to handle a windows legacy ...
						encoding = RtfSpec.AnsiEncoding;
						break;
					default:
						encoding = Encoding.GetEncoding( codePage );
						break;
				}
				byteToCharDecoder = null;
			}
			if ( byteToCharDecoder == null )
			{
				byteToCharDecoder = encoding.GetDecoder();
			}
		} // UpdateEncoding

		// ----------------------------------------------------------------------
		private static bool IsASCIILetter( int character )
		{
			return (character >= 'a' && character <= 'z') || (character >= 'A' && character <= 'Z');
		} // IsASCIILetter

		// ----------------------------------------------------------------------
		private static bool IsHexDigit( int character )
		{
			return (character >= '0' && character <= '9') ||
						 (character >= 'a' && character <= 'f') ||
						 (character >= 'A' && character <= 'F');
		} // IsHexDigit

		// ----------------------------------------------------------------------
		private static bool IsDigit( int character )
		{
			return character >= '0' && character <= '9';
		} // IsDigit

		// ----------------------------------------------------------------------
		private static int ReadOneByte( TextReader reader )
		{
			int byteValue = reader.Read();
			if ( byteValue == -1 )
			{
				throw new RtfUnicodeEncodingException( Strings.UnexpectedEndOfFile );
			}
			return byteValue;
		} // ReadOneByte

		// ----------------------------------------------------------------------
		private char ReadOneChar( TextReader reader )
		{
			// NOTE: the handling of multi-byte encodings is probably not the most
			// efficient here ...

			bool completed = false;
			int byteIndex = 0;
			while ( !completed )
			{
				byteDecodingBuffer[ byteIndex ] = (byte)ReadOneByte( reader );
				byteIndex++;
				int usedBytes;
				int usedChars;
				byteToCharDecoder.Convert(
					byteDecodingBuffer, 0, byteIndex,
					charDecodingBuffer, 0, 1,
					true,
					out usedBytes,
					out usedChars,
					out completed );
				if ( completed && ( usedBytes != byteIndex || usedChars != 1 ) )
				{
					throw new RtfMultiByteEncodingException( Strings.InvalidMultiByteEncoding( 
					byteDecodingBuffer, byteIndex, encoding ) );
				}
			}
			char character = charDecodingBuffer[ 0 ];
			return character;
		} // ReadOneChar

		// ----------------------------------------------------------------------
		private void DecodeCurrentHexBuffer()
		{
			long pendingByteCount = hexDecodingBuffer.Length;
			if ( pendingByteCount > 0 )
			{
				byte[] pendingBytes = hexDecodingBuffer.ToArray();
				char[] convertedChars = new char[ pendingByteCount ]; // should be enough

				int startIndex = 0;
				bool completed = false;
				while ( !completed && startIndex < pendingBytes.Length )
				{
					int usedBytes;
					int usedChars;
					byteToCharDecoder.Convert(
						pendingBytes, startIndex, pendingBytes.Length - startIndex,
						convertedChars, 0, convertedChars.Length,
						true,
						out usedBytes,
						out usedChars,
						out completed );
					curText.Append( convertedChars, 0, usedChars );
					startIndex += usedChars;
				}

				hexDecodingBuffer.SetLength( 0 );
			}
		} // DecodeCurrentHexBuffer

		// ----------------------------------------------------------------------
// ReSharper disable UnusedParameter.Local
		private static int PeekNextChar( TextReader reader, bool mandatory )
// ReSharper restore UnusedParameter.Local
		{
			int character = reader.Peek();
			if ( mandatory && character == -1 )
			{
				throw new RtfMultiByteEncodingException( Strings.EndOfFileInvalidCharacter );
			}
			return character;
		} // PeekNextChar

		// ----------------------------------------------------------------------
		private void FlushText()
		{
			if ( curText.Length > 0 )
			{
				if ( level == 0 )
				{
					throw new RtfStructureException( Strings.TextOnRootLevel( curText.ToString() ) );
				}
				NotifyTextFound( new RtfText( curText.ToString() ) );
				curText.Remove( 0, curText.Length );
			}
		} // FlushText

		// ----------------------------------------------------------------------
		// members
		private StringBuilder curText;
		private readonly Stack unicodeSkipCountStack = new Stack();
		private int unicodeSkipCount;
		private readonly Stack codePageStack = new Stack();
		private int level;
		private int tagCountAtLastGroupStart;
		private int tagCount;
		private int fontTableStartLevel;
		private string targetFont;
		private bool expectingThemeFont;
		private readonly Hashtable fontToCodePageMapping = new Hashtable();
		private Encoding encoding;
		private Decoder byteToCharDecoder;
		private readonly MemoryStream hexDecodingBuffer = new MemoryStream();
		private readonly byte[] byteDecodingBuffer = new byte[ 8 ]; // >0 for multi-byte encodings
		private readonly char[] charDecodingBuffer = new char[ 1 ];

	} // class RtfParser

} // namespace Itenso.Rtf.Parser
// -- EOF -------------------------------------------------------------------
