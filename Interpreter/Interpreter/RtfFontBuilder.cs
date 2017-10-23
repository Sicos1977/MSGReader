// -- FILE ------------------------------------------------------------------
// name       : RtfFontBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System.Text;
using Itenso.Rtf.Model;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{

	// ------------------------------------------------------------------------
	public sealed class RtfFontBuilder : RtfElementVisitorBase
	{

		// ----------------------------------------------------------------------
		public RtfFontBuilder() :
			base( RtfElementVisitorOrder.NonRecursive )
		{
			// we iterate over our children ourselves -> hence non-recursive
			Reset();
		} // RtfFontBuilder

		// ----------------------------------------------------------------------
		public string FontId
		{
			get { return fontId; }
		} // FontId

		// ----------------------------------------------------------------------
		public int FontIndex
		{
			get { return fontIndex; }
		} // FontIndex

		// ----------------------------------------------------------------------
		public int FontCharset
		{
			get { return fontCharset; }
		} // FontCharset

		// ----------------------------------------------------------------------
		public int FontCodePage
		{
			get { return fontCodePage; }
		} // FontCodePage

		// ----------------------------------------------------------------------
		public RtfFontKind FontKind
		{
			get { return fontKind; }
		} // FontKind

		// ----------------------------------------------------------------------
		public RtfFontPitch FontPitch
		{
			get { return fontPitch; }
		} // FontPitch

		// ----------------------------------------------------------------------
		public string FontName
		{
			get
			{
				string fontName = null;
				int len = fontNameBuffer.Length;
				if ( len > 0 && fontNameBuffer[ len - 1 ] == ';' )
				{
					fontName = fontNameBuffer.ToString().Substring( 0, len - 1 ).Trim();
					if ( fontName.Length == 0 )
					{
						fontName = null;
					}
				}
				return fontName;
			}
		} // FontName

		// ----------------------------------------------------------------------
		public IRtfFont CreateFont()
		{
			string fontName = FontName;
			if ( string.IsNullOrEmpty( fontName ) )
			{
				fontName = "UnnamedFont_" + fontId;
			}
			return new RtfFont( fontId, fontKind, fontPitch,
				fontCharset, fontCodePage, fontName );
		} // CreateFont

		// ----------------------------------------------------------------------
		public void Reset()
		{
			fontIndex = 0;
			fontCharset = 0;
			fontCodePage = 0;
			fontKind = RtfFontKind.Nil;
			fontPitch = RtfFontPitch.Default;
			fontNameBuffer.Remove( 0, fontNameBuffer.Length );
		} // Reset

		// ----------------------------------------------------------------------
		protected override void DoVisitGroup( IRtfGroup group )
		{
			switch ( group.Destination )
			{
				case RtfSpec.TagFont:
				case RtfSpec.TagThemeFontLoMajor:
				case RtfSpec.TagThemeFontHiMajor:
				case RtfSpec.TagThemeFontDbMajor:
				case RtfSpec.TagThemeFontBiMajor:
				case RtfSpec.TagThemeFontLoMinor:
				case RtfSpec.TagThemeFontHiMinor:
				case RtfSpec.TagThemeFontDbMinor:
				case RtfSpec.TagThemeFontBiMinor:
					VisitGroupChildren( group );
					break;
			}
		} // DoVisitGroup

		// ----------------------------------------------------------------------
		protected override void DoVisitTag( IRtfTag tag )
		{
			switch ( tag.Name )
			{
				case RtfSpec.TagThemeFontLoMajor:
				case RtfSpec.TagThemeFontHiMajor:
				case RtfSpec.TagThemeFontDbMajor:
				case RtfSpec.TagThemeFontBiMajor:
				case RtfSpec.TagThemeFontLoMinor:
				case RtfSpec.TagThemeFontHiMinor:
				case RtfSpec.TagThemeFontDbMinor:
				case RtfSpec.TagThemeFontBiMinor:
					// skip and ignore for the moment
					break;
				case RtfSpec.TagFont:
					fontId = tag.FullName;
					fontIndex = tag.ValueAsNumber;
					break;
				case RtfSpec.TagFontKindNil:
					fontKind = RtfFontKind.Nil;
					break;
				case RtfSpec.TagFontKindRoman:
					fontKind = RtfFontKind.Roman;
					break;
				case RtfSpec.TagFontKindSwiss:
					fontKind = RtfFontKind.Swiss;
					break;
				case RtfSpec.TagFontKindModern:
					fontKind = RtfFontKind.Modern;
					break;
				case RtfSpec.TagFontKindScript:
					fontKind = RtfFontKind.Script;
					break;
				case RtfSpec.TagFontKindDecor:
					fontKind = RtfFontKind.Decor;
					break;
				case RtfSpec.TagFontKindTech:
					fontKind = RtfFontKind.Tech;
					break;
				case RtfSpec.TagFontKindBidi:
					fontKind = RtfFontKind.Bidi;
					break;
				case RtfSpec.TagFontCharset:
					fontCharset = tag.ValueAsNumber;
					break;
				case RtfSpec.TagCodePage:
					fontCodePage = tag.ValueAsNumber;
					break;
				case RtfSpec.TagFontPitch:
					switch ( tag.ValueAsNumber )
					{
						case 0:
							fontPitch = RtfFontPitch.Default;
							break;
						case 1:
							fontPitch = RtfFontPitch.Fixed;
							break;
						case 2:
							fontPitch = RtfFontPitch.Variable;
							break;
					}
					break;
			}
		} // DoVisitTag

		// ----------------------------------------------------------------------
		protected override void DoVisitText( IRtfText text )
		{
			fontNameBuffer.Append( text.Text );
		} // DoVisitText

		// ----------------------------------------------------------------------
		// members
		private string fontId;
		private int fontIndex;
		private int fontCharset;
		private int fontCodePage;
		private RtfFontKind fontKind;
		private RtfFontPitch fontPitch;
		private readonly StringBuilder fontNameBuffer = new StringBuilder();

	} // class RtfFontBuilder

} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------
