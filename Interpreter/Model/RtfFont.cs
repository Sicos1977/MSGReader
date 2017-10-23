// -- FILE ------------------------------------------------------------------
// name       : RtfFont.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfFont : IRtfFont
	{

		// ----------------------------------------------------------------------
		public RtfFont( string id, RtfFontKind kind, RtfFontPitch pitch, int charSet, int codePage, string name )
		{
			if ( id == null )
			{
				throw new ArgumentNullException( "id" );
			}
			if ( charSet < 0 )
			{
				throw new ArgumentException( Strings.InvalidCharacterSet( charSet ) );
			}
			if ( codePage < 0 )
			{
				throw new ArgumentException( Strings.InvalidCodePage( codePage ) );
			}
			if ( name == null )
			{
				throw new ArgumentNullException( "name" );
			}
			this.id = id;
			this.kind = kind;
			this.pitch = pitch;
			this.charSet = charSet;
			this.codePage = codePage;
			this.name = name;
		} // RtfFont

		// ----------------------------------------------------------------------
		public string Id
		{
			get { return id; }
		} // Id

		// ----------------------------------------------------------------------
		public RtfFontKind Kind
		{
			get { return kind; }
		} // Kind

		// ----------------------------------------------------------------------
		public RtfFontPitch Pitch
		{
			get { return pitch; }
		} // Pitch

		// ----------------------------------------------------------------------
		public int CharSet
		{
			get { return charSet; }
		} // CharSet

		// ----------------------------------------------------------------------
		public int CodePage
		{
			get
			{
				// if a codepage is specified, it overrides the charset setting
				if ( codePage == 0 )
				{
					// unspecified codepage: use the one derived from the charset:
					return RtfSpec.GetCodePage( charSet );
				}
				return codePage;
			}
		} // CodePage

		// ----------------------------------------------------------------------
		public Encoding GetEncoding()
		{
			return Encoding.GetEncoding( CodePage );
		} // GetEncoding

		// ----------------------------------------------------------------------
		public string Name
		{
			get { return name; }
		} // Name

		// ----------------------------------------------------------------------
		public override bool Equals( object obj )
		{
			if ( obj == this )
			{
				return true;
			}

			if ( obj == null || GetType() != obj.GetType() )
			{
				return false;
			}

			return IsEqual( obj );
		} // Equals

		// ----------------------------------------------------------------------
		public override int GetHashCode()
		{
			return HashTool.AddHashCode( GetType().GetHashCode(), ComputeHashCode() );
		} // GetHashCode

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return id + ":" + name;
		} // ToString

		// ----------------------------------------------------------------------
		private bool IsEqual( object obj )
		{
			RtfFont compare = obj as RtfFont; // guaranteed to be non-null
			return
				compare != null &&
				id.Equals( compare.id ) &&
				kind == compare.kind &&
				pitch == compare.pitch &&
				charSet == compare.charSet &&
				codePage == compare.codePage &&
				name.Equals( compare.name );
		} // IsEqual

		// ----------------------------------------------------------------------
		private int ComputeHashCode()
		{
			int hash = id.GetHashCode();
			hash = HashTool.AddHashCode( hash, kind );
			hash = HashTool.AddHashCode( hash, pitch );
			hash = HashTool.AddHashCode( hash, charSet );
			hash = HashTool.AddHashCode( hash, codePage );
			hash = HashTool.AddHashCode( hash, name );
			return hash;
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		// members
		private readonly string id;
		private readonly RtfFontKind kind;
		private readonly RtfFontPitch pitch;
		private readonly int charSet;
		private readonly int codePage;
		private readonly string name;

	} // class RtfFont

} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------
