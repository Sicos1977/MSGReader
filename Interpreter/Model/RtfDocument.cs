// -- FILE ------------------------------------------------------------------
// name       : RtfDocument.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfDocument : IRtfDocument
	{

		// ----------------------------------------------------------------------
		public RtfDocument( IRtfInterpreterContext context, IRtfVisualCollection visualContent ) :
			this( context.RtfVersion,
				context.DefaultFont,
				context.FontTable,
				context.ColorTable,
				context.Generator,
				context.UniqueTextFormats,
				context.DocumentInfo,
				context.UserProperties,
				visualContent
			)
		{
		} // RtfDocument

		// ----------------------------------------------------------------------
		public RtfDocument(
			int rtfVersion,
			IRtfFont defaultFont,
			IRtfFontCollection fontTable,
			IRtfColorCollection colorTable,
			string generator,
			IRtfTextFormatCollection uniqueTextFormats,
			IRtfDocumentInfo documentInfo,
			IRtfDocumentPropertyCollection userProperties,
			IRtfVisualCollection visualContent
		)
		{
			if ( rtfVersion != RtfSpec.RtfVersion1 )
			{
				throw new RtfUnsupportedStructureException( Strings.UnsupportedRtfVersion( rtfVersion ) );
			}
			if ( defaultFont == null )
			{
				throw new ArgumentNullException( "defaultFont" );
			}
			if ( fontTable == null )
			{
				throw new ArgumentNullException( "fontTable" );
			}
			if ( colorTable == null )
			{
				throw new ArgumentNullException( "colorTable" );
			}
			if ( uniqueTextFormats == null )
			{
				throw new ArgumentNullException( "uniqueTextFormats" );
			}
			if ( documentInfo == null )
			{
				throw new ArgumentNullException( "documentInfo" );
			}
			if ( userProperties == null )
			{
				throw new ArgumentNullException( "userProperties" );
			}
			if ( visualContent == null )
			{
				throw new ArgumentNullException( "visualContent" );
			}
			this.rtfVersion = rtfVersion;
			this.defaultFont = defaultFont;
			defaultTextFormat = new RtfTextFormat( defaultFont, RtfSpec.DefaultFontSize );
			this.fontTable = fontTable;
			this.colorTable = colorTable;
			this.generator = generator;
			this.uniqueTextFormats = uniqueTextFormats;
			this.documentInfo = documentInfo;
			this.userProperties = userProperties;
			this.visualContent = visualContent;
		} // RtfDocument

		// ----------------------------------------------------------------------
		public int RtfVersion
		{
			get { return rtfVersion; }
		} // RtfVersion

		// ----------------------------------------------------------------------
		public IRtfFont DefaultFont
		{
			get { return defaultFont; }
		} // DefaultFont

		// ----------------------------------------------------------------------
		public IRtfTextFormat DefaultTextFormat
		{
			get { return defaultTextFormat; }
		} // DefaultTextFormat

		// ----------------------------------------------------------------------
		public IRtfFontCollection FontTable
		{
			get { return fontTable; }
		} // FontTable

		// ----------------------------------------------------------------------
		public IRtfColorCollection ColorTable
		{
			get { return colorTable; }
		} // ColorTable

		// ----------------------------------------------------------------------
		public string Generator
		{
			get { return generator; }
		} // Generator

		// ----------------------------------------------------------------------
		public IRtfTextFormatCollection UniqueTextFormats
		{
			get { return uniqueTextFormats; }
		} // UniqueTextFormats

		// ----------------------------------------------------------------------
		public IRtfDocumentInfo DocumentInfo
		{
			get { return documentInfo; }
		} // DocumentInfo

		// ----------------------------------------------------------------------
		public IRtfDocumentPropertyCollection UserProperties
		{
			get { return userProperties; }
		} // UserProperties

		// ----------------------------------------------------------------------
		public IRtfVisualCollection VisualContent
		{
			get { return visualContent; }
		} // VisualContent

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return "RTFv" + rtfVersion;
		} // ToString

		// ----------------------------------------------------------------------
		// members
		private readonly int rtfVersion;
		private readonly IRtfFont defaultFont;
		private readonly IRtfTextFormat defaultTextFormat;
		private readonly IRtfFontCollection fontTable;
		private readonly IRtfColorCollection colorTable;
		private readonly string generator;
		private readonly IRtfTextFormatCollection uniqueTextFormats;
		private readonly IRtfDocumentInfo documentInfo;
		private readonly IRtfDocumentPropertyCollection userProperties;
		private readonly IRtfVisualCollection visualContent;

	} // class RtfDocument

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
