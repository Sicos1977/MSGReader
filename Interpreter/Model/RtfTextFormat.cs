// -- FILE ------------------------------------------------------------------
// name       : IRtfTextFormat.cs
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
	public sealed class RtfTextFormat : IRtfTextFormat
	{

		// ----------------------------------------------------------------------
		public RtfTextFormat( IRtfFont font, int fontSize )
		{
			if ( font == null )
			{
				throw new ArgumentNullException( "font" );
			}
			if ( fontSize <= 0 || fontSize > 0xFFFF )
			{
				throw new ArgumentException( Strings.FontSizeOutOfRange( fontSize ) );
			}
			this.font = font;
			this.fontSize = fontSize;
		} // RtfTextFormat

		// ----------------------------------------------------------------------
		public RtfTextFormat( IRtfTextFormat copy )
		{
			if ( copy == null )
			{
				throw new ArgumentNullException( "copy" );
			}
			font = copy.Font; // enough because immutable
			fontSize = copy.FontSize;
			superScript = copy.SuperScript;
			bold = copy.IsBold;
			italic = copy.IsItalic;
			underline = copy.IsUnderline;
			strikeThrough = copy.IsStrikeThrough;
			hidden = copy.IsHidden;
			backgroundColor = copy.BackgroundColor; // enough because immutable
			foregroundColor = copy.ForegroundColor; // enough because immutable
			alignment = copy.Alignment;
		} // RtfTextFormat

		// ----------------------------------------------------------------------
		public RtfTextFormat( RtfTextFormat copy )
		{
			if ( copy == null )
			{
				throw new ArgumentNullException( "copy" );
			}
			font = copy.font; // enough because immutable
			fontSize = copy.fontSize;
			superScript = copy.superScript;
			bold = copy.bold;
			italic = copy.italic;
			underline = copy.underline;
			strikeThrough = copy.strikeThrough;
			hidden = copy.hidden;
			backgroundColor = copy.backgroundColor; // enough because immutable
			foregroundColor = copy.foregroundColor; // enough because immutable
			alignment = copy.alignment;
		} // RtfTextFormat

		// ----------------------------------------------------------------------
		public string FontDescriptionDebug
		{
			get
			{
				StringBuilder buf = new StringBuilder( font.Name );
				buf.Append( ", " );
				buf.Append( fontSize );
				buf.Append( superScript >= 0 ? "+" : "" );
				buf.Append( superScript );
				buf.Append( ", " );
				if ( bold || italic || underline || strikeThrough )
				{
					bool combined = false;
					if ( bold )
					{
						buf.Append( "bold" );
						combined = true;
					}
					if ( italic )
					{
						buf.Append( combined ? "+italic" : "italic" );
						combined = true;
					}
					if ( underline )
					{
						buf.Append( combined ? "+underline" : "underline" );
						combined = true;
					}
					if ( strikeThrough )
					{
						buf.Append( combined ? "+strikethrough" : "strikethrough" );
					}
				}
				else
				{
					buf.Append( "plain" );
				}
				if ( hidden )
				{
					buf.Append( ", hidden" );
				}
				return buf.ToString();
			}
		} // FontDescriptionDebug

		// ----------------------------------------------------------------------
		public IRtfFont Font
		{
			get { return font; }
		} // Font

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithFont( IRtfFont rtfFont )
		{
			if ( rtfFont == null )
			{
				throw new ArgumentNullException( "rtfFont" );
			}
			if ( font.Equals( rtfFont ) )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.font = rtfFont;
			return copy;
		} // DeriveWithFont

		// ----------------------------------------------------------------------
		public int FontSize
		{
			get { return fontSize; }
		} // FontSize

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithFontSize( int derivedFontSize )
		{
			if ( derivedFontSize < 0 || derivedFontSize > 0xFFFF )
			{
				throw new ArgumentException( Strings.FontSizeOutOfRange( derivedFontSize ) );
			}
			if ( fontSize == derivedFontSize )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.fontSize = derivedFontSize;
			return copy;
		} // DeriveWithFontSize

		// ----------------------------------------------------------------------
		public int SuperScript
		{
			get { return superScript; }
		} // SuperScript

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithSuperScript( int deviation )
		{
			if ( superScript == deviation )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.superScript = deviation;
			// reset font size
			if ( deviation == 0 )
			{
				copy.fontSize = ( fontSize / 2 ) * 3;
			}
			return copy;
		} // DeriveWithSuperScript

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithSuperScript( bool super )
		{
			RtfTextFormat copy = new RtfTextFormat( this );
			copy.fontSize = Math.Max( 1, ( fontSize * 2 ) / 3 );
			copy.superScript = ( super ? 1 : -1 ) * Math.Max( 1, fontSize / 2 );
			return copy;
		} // DeriveWithSuperScript

		// ----------------------------------------------------------------------
		public bool IsNormal
		{
			get
			{
				return
					!bold && !italic && !underline && !strikeThrough &&
					!hidden &&
					fontSize == RtfSpec.DefaultFontSize &&
					superScript == 0 &&
					RtfColor.Black.Equals( foregroundColor ) &&
					RtfColor.White.Equals( backgroundColor );
			}
		} // IsNormal

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveNormal()
		{
			if ( IsNormal )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( font, RtfSpec.DefaultFontSize );
			copy.alignment = alignment; // this is a paragraph property, keep it
			return copy;
		} // DeriveNormal

		// ----------------------------------------------------------------------
		public bool IsBold
		{
			get { return bold; }
		} // IsBold

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithBold( bool derivedBold )
		{
			if ( bold == derivedBold )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.bold = derivedBold;
			return copy;
		} // DeriveWithBold

		// ----------------------------------------------------------------------
		public bool IsItalic
		{
			get { return italic; }
		} // IsItalic

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithItalic( bool derivedItalic )
		{
			if ( italic == derivedItalic )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.italic = derivedItalic;
			return copy;
		} // DeriveWithItalic

		// ----------------------------------------------------------------------
		public bool IsUnderline
		{
			get { return underline; }
		} // IsUnderline

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithUnderline( bool derivedUnderline )
		{
			if ( underline == derivedUnderline )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.underline = derivedUnderline;
			return copy;
		} // DeriveWithUnderline

		// ----------------------------------------------------------------------
		public bool IsStrikeThrough
		{
			get { return strikeThrough; }
		} // IsStrikeThrough

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithStrikeThrough( bool derivedStrikeThrough )
		{
			if ( strikeThrough == derivedStrikeThrough )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.strikeThrough = derivedStrikeThrough;
			return copy;
		} // DeriveWithStrikeThrough

		// ----------------------------------------------------------------------
		public bool IsHidden
		{
			get { return hidden; }
		} // IsHidden

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithHidden( bool derivedHidden )
		{
			if ( hidden == derivedHidden )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.hidden = derivedHidden;
			return copy;
		} // DeriveWithHidden

		// ----------------------------------------------------------------------
		public IRtfColor BackgroundColor
		{
			get { return backgroundColor; }
		} // BackgroundColor

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithBackgroundColor( IRtfColor derivedBackgroundColor )
		{
			if ( derivedBackgroundColor == null )
			{
				throw new ArgumentNullException( "derivedBackgroundColor" );
			}
			if ( backgroundColor.Equals( derivedBackgroundColor ) )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.backgroundColor = derivedBackgroundColor;
			return copy;
		} // DeriveWithBackgroundColor

		// ----------------------------------------------------------------------
		public IRtfColor ForegroundColor
		{
			get { return foregroundColor; }
		} // ForegroundColor

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithForegroundColor( IRtfColor derivedForegroundColor )
		{
			if ( derivedForegroundColor == null )
			{
				throw new ArgumentNullException( "derivedForegroundColor" );
			}
			if ( foregroundColor.Equals( derivedForegroundColor ) )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.foregroundColor = derivedForegroundColor;
			return copy;
		} // DeriveWithForegroundColor

		// ----------------------------------------------------------------------
		public RtfTextAlignment Alignment
		{
			get { return alignment; }
		} // Alignment

		// ----------------------------------------------------------------------
		public RtfTextFormat DeriveWithAlignment( RtfTextAlignment derivedAlignment )
		{
			if ( alignment == derivedAlignment )
			{
				return this;
			}

			RtfTextFormat copy = new RtfTextFormat( this );
			copy.alignment = derivedAlignment;
			return copy;
		} // DeriveWithForegroundColor

		// ----------------------------------------------------------------------
		IRtfTextFormat IRtfTextFormat.Duplicate()
		{
			return new RtfTextFormat( this );
		} // IRtfTextFormat.Duplicate

		// ----------------------------------------------------------------------
		public RtfTextFormat Duplicate()
		{
			return new RtfTextFormat( this );
		} // Duplicate

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
		private bool IsEqual( object obj )
		{
			RtfTextFormat compare = obj as RtfTextFormat; // guaranteed to be non-null
			return
				compare != null &&
				font.Equals( compare.font ) &&
				fontSize == compare.fontSize &&
				superScript == compare.superScript &&
				bold == compare.bold &&
				italic == compare.italic &&
				underline == compare.underline &&
				strikeThrough == compare.strikeThrough &&
				hidden == compare.hidden &&
				backgroundColor.Equals( compare.backgroundColor ) &&
				foregroundColor.Equals( compare.foregroundColor ) &&
				alignment == compare.alignment;
		} // IsEqual

		// ----------------------------------------------------------------------
		private int ComputeHashCode()
		{
			int hash = font.GetHashCode();
			hash = HashTool.AddHashCode( hash, fontSize );
			hash = HashTool.AddHashCode( hash, superScript );
			hash = HashTool.AddHashCode( hash, bold );
			hash = HashTool.AddHashCode( hash, italic );
			hash = HashTool.AddHashCode( hash, underline );
			hash = HashTool.AddHashCode( hash, strikeThrough );
			hash = HashTool.AddHashCode( hash, hidden );
			hash = HashTool.AddHashCode( hash, backgroundColor );
			hash = HashTool.AddHashCode( hash, foregroundColor );
			hash = HashTool.AddHashCode( hash, alignment );
			return hash;
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			StringBuilder buf = new StringBuilder( "Font " );
			buf.Append( FontDescriptionDebug );
			buf.Append( ", " );
			buf.Append( alignment );
			buf.Append( ", " );
			buf.Append( foregroundColor.ToString() );
			buf.Append( " on " );
			buf.Append( backgroundColor.ToString() );
			return buf.ToString();
		} // ToString

		// ----------------------------------------------------------------------
		// members
		private IRtfFont font;
		private int fontSize;
		private int superScript;
		private bool bold;
		private bool italic;
		private bool underline;
		private bool strikeThrough;
		private bool hidden;
		private IRtfColor backgroundColor = RtfColor.White;
		private IRtfColor foregroundColor = RtfColor.Black;
		private RtfTextAlignment alignment = RtfTextAlignment.Left;

	} // class RtfTextFormat

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
