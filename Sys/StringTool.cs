// -- FILE ------------------------------------------------------------------
// name       : StringTool.cs
// project    : System Framelet
// created    : Leon Poyyayil - 2005.11.28
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Globalization;
using System.Text;
using Itenso.Sys.Collection;

namespace Itenso.Sys
{

	// ------------------------------------------------------------------------
	public static class StringTool
	{

		// ----------------------------------------------------------------------
		public static string FormatSafeInvariant( string format, params object[] args )
		{
			try
			{
				return string.Format( CultureInfo.InvariantCulture, format, args );
			}
			catch (FormatException)
			{
				return format;
			}
		} // FormatSafeInvariant

		// ----------------------------------------------------------------------
		/// <summary>
		/// Splits a string in the same way as the System.String.Split() method but
		/// with support for special treatment for escaped characters and for quoted
		/// sections which won't be split.
		/// </summary>
		/// <remarks>
		/// Escaping supports the following special treatment:
		/// <list type="bullet">
		/// <item>escape is followed by 'n': a new line character is inserted</item>
		/// <item>escape is followed by 'r': a form feed character is inserted</item>
		/// <item>escape is followed by 't': a tabulator character is inserted</item>
		/// <item>escape is followed by 'x': the next two characters are interpreted
		/// as a hex code of the character to be inserted</item>
		/// <item>any other character after the escape is inserted literally</item>
		/// </list>
		/// Escaping is applied within and outside of quoted sections.
		/// </remarks>
		/// <param name="toSplit">the string to split</param>
		/// <param name="quote">the quoting character, e.g. '&quot;'</param>
		/// <param name="escape">an escaping character to use both within and outside
		/// of quoted sections, e.g. '\\'</param>
		/// <param name="includeEmptyUnquotedSections">whether to return zero-length sections
		/// outside of quotations or not. empty sections adjacent to a quoted section are
		/// never returned.</param>
		/// <param name="separator">the separator character(s) which will be used to
		/// split the string outside of quoted sections</param>
		/// <returns>the array of sections into which the string has been split up.
		/// never null but possibly empty.</returns>
		public static string[] SplitQuoted( string toSplit, char quote, char escape,
			bool includeEmptyUnquotedSections, params char[] separator )
		{
			if ( toSplit == null )
			{
				throw new ArgumentNullException( "toSplit" );
			}
			if ( separator == null || separator.Length == 0 )
			{
				throw new ArgumentNullException( "separator" );
			}
			string separators = new string( separator );
			if ( separators.IndexOf( quote ) >= 0 || separators.IndexOf( escape ) >= 0 )
			{
				throw new ArgumentException( Strings.StringToolSeparatorIncludesQuoteOrEscapeChar, "separator" );
			}

			StringCollection sections = new StringCollection();

			StringBuilder section = null;
			int length = toSplit.Length;
			bool inQuotedSection = false;
			for ( int i = 0; i < length; i++ )
			{
				char c = toSplit[ i ];
				if ( c == escape )
				{
					if ( i < length - 1 )
					{
						if ( section == null )
						{
							section = new StringBuilder();
						}
						i++;
						c = toSplit[ i ];
						switch ( c )
						{
							case 'n':
								section.Append( '\n' );
								break;
							case 'r':
								section.Append( '\r' );
								break;
							case 't':
								section.Append( '\t' );
								break;
							case 'x':
								if ( i < length - 2 )
								{
									int upperHexNibble = GetHexValue( toSplit[ i + 1 ] ) * 16;
									int lowerHexNibble = GetHexValue( toSplit[ i + 2 ] );
									char hexChar = (char)( upperHexNibble + lowerHexNibble );
									section.Append( hexChar );
									i += 2;
								}
								else
								{
									throw new ArgumentException( Strings.StringToolMissingEscapedHexCode, "toSplit" );
								}
								break;
							default:
								section.Append( c );
								break;
						}
					}
					else
					{
						throw new ArgumentException( Strings.StringToolMissingEscapedChar, "toSplit" );
					}
				}
				else if ( c == quote )
				{
					if ( section != null )
					{
						sections.Add( section.ToString() );
						section = null;
					}
					else if ( inQuotedSection )
					{
						sections.Add( string.Empty );
					}
					inQuotedSection = !inQuotedSection;
				}
				else if ( separators.IndexOf( c ) >= 0 )
				{
					if ( inQuotedSection )
					{
						if ( section == null )
						{
							section = new StringBuilder();
						}
						section.Append( c );
					}
					else
					{
						if ( section != null )
						{
							sections.Add( section.ToString() );
							section = null;
						}
						else if ( includeEmptyUnquotedSections )
						{
							if ( i == 0 || separators.IndexOf( toSplit[ i - 1 ] ) >= 0 )
							{
								sections.Add( string.Empty );
							}
						}
					}
				}
				else
				{
					if ( section == null )
					{
						section = new StringBuilder();
					}
					section.Append( c );
				}
			}
			if ( inQuotedSection )
			{
				throw new ArgumentException( Strings.StringToolUnbalancedQuotes, "toSplit" );
			}
			if ( section != null )
			{
				sections.Add( section.ToString() );
			}

			string[] sectionArray = new string[ sections.Count ];
			sections.CopyTo( sectionArray, 0 );
			return sectionArray;
		} // SplitQuoted

		// ----------------------------------------------------------------------
		private static int GetHexValue( char c )
		{
			if ( c >= 'a' && c <= 'f' )
			{
				return c - 'a' + 10;
			}

			if ( c >= 'A' && c <= 'F' )
			{
				return c - 'A' + 10;
			}

			if ( c >= '0' && c <= '9' )
			{
				return c - '0';
			}
			throw new ArgumentException( Strings.StringToolContainsInvalidHexChar, "c" );
		} // GetHexValue

		// ----------------------------------------------------------------------
		// members

	} // class StringTool

} // namespace Itenso.Sys
// -- EOF -------------------------------------------------------------------
