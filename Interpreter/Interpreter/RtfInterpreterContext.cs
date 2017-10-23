// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreterContext.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Collections;
using Itenso.Rtf.Model;

namespace Itenso.Rtf.Interpreter
{

	// ------------------------------------------------------------------------
	public sealed class RtfInterpreterContext : IRtfInterpreterContext
	{

		// ----------------------------------------------------------------------
		public RtfInterpreterState State
		{
			get { return state; }
			set { state = value; }
		} // State

		// ----------------------------------------------------------------------
		public int RtfVersion
		{
			get { return rtfVersion; }
			set { rtfVersion = value; }
		} // RtfVersion

		// ----------------------------------------------------------------------
		public string DefaultFontId
		{
			get { return defaultFontId; }
			set { defaultFontId = value; }
		} // DefaultFontIndex

		// ----------------------------------------------------------------------
		public IRtfFont DefaultFont
		{
			get
			{
				IRtfFont defaultFont = fontTable[ defaultFontId ];
				if ( defaultFont != null )
				{
					return defaultFont;
				}
				throw new RtfUndefinedFontException( Strings.InvalidDefaultFont( 
					defaultFontId, fontTable.ToString() ) );
			}
		} // DefaultFont

		// ----------------------------------------------------------------------
		public IRtfFontCollection FontTable
		{
			get { return fontTable; }
		} // FontTable

		// ----------------------------------------------------------------------
		public RtfFontCollection WritableFontTable
		{
			get { return fontTable; }
		} // WritableFontTable

		// ----------------------------------------------------------------------
		public IRtfColorCollection ColorTable
		{
			get { return colorTable; }
		} // ColorTable

		// ----------------------------------------------------------------------
		public RtfColorCollection WritableColorTable
		{
			get { return colorTable; }
		} // WritableColorTable

		// ----------------------------------------------------------------------
		public string Generator
		{
			get { return generator; }
			set { generator = value; }
		} // Generator

		// ----------------------------------------------------------------------
		public IRtfTextFormatCollection UniqueTextFormats
		{
			get { return uniqueTextFormats; }
		} // UniqueTextFormats

		// ----------------------------------------------------------------------
		public IRtfTextFormat CurrentTextFormat
		{
			get { return currentTextFormat; }
		} // CurrentTextFormat

		// ----------------------------------------------------------------------
		public IRtfTextFormat GetSafeCurrentTextFormat()
		{
			return currentTextFormat ?? WritableCurrentTextFormat;
		} // GetSafeCurrentTextFormat

		// ----------------------------------------------------------------------
		public IRtfTextFormat GetUniqueTextFormatInstance( IRtfTextFormat templateFormat )
		{
			if ( templateFormat == null )
			{
				throw new ArgumentNullException( "templateFormat" );
			}
			IRtfTextFormat uniqueInstance;
			int existingEquivalentPos = uniqueTextFormats.IndexOf( templateFormat );
			if ( existingEquivalentPos >= 0 )
			{
				// we already know an equivalent format -> reference that one for future use
				uniqueInstance = uniqueTextFormats[ existingEquivalentPos ];
			}
			else
			{
				// this is a yet unknown format -> add it to the known formats and use it
				uniqueTextFormats.Add( templateFormat );
				uniqueInstance = templateFormat;
			}
			return uniqueInstance;
		} // GetUniqueTextFormatInstance

		// ----------------------------------------------------------------------
		public RtfTextFormat WritableCurrentTextFormat
		{
			get
			{
				if ( currentTextFormat == null )
				{
					// set via property to ensure it will get added to the unique map
					WritableCurrentTextFormat = new RtfTextFormat( DefaultFont, RtfSpec.DefaultFontSize );
				}
				return currentTextFormat;
			}
			set
			{
				currentTextFormat = (RtfTextFormat)GetUniqueTextFormatInstance( value );
			}
		} // WritableCurrentTextFormat

		// ----------------------------------------------------------------------
		public IRtfDocumentInfo DocumentInfo
		{
			get { return documentInfo; }
		} // DocumentInfo

		// ----------------------------------------------------------------------
		public RtfDocumentInfo WritableDocumentInfo
		{
			get { return documentInfo; }
		} // WritableDocumentInfo

		// ----------------------------------------------------------------------
		public IRtfDocumentPropertyCollection UserProperties
		{
			get { return userProperties; }
		} // UserProperties

		// ----------------------------------------------------------------------
		public RtfDocumentPropertyCollection WritableUserProperties
		{
			get { return userProperties; }
		} // WritableUserProperties

		// ----------------------------------------------------------------------
		public void PushCurrentTextFormat()
		{
			textFormatStack.Push( WritableCurrentTextFormat );
		} // PushCurrentTextFormat

		// ----------------------------------------------------------------------
		public void PopCurrentTextFormat()
		{
			if ( textFormatStack.Count == 0 )
			{
				throw new RtfStructureException( Strings.InvalidTextContextState );
			}
			currentTextFormat = (RtfTextFormat)textFormatStack.Pop();
		} // PopCurrentTextFormat

		// ----------------------------------------------------------------------
		public void Reset()
		{
			state = RtfInterpreterState.Init;
			rtfVersion = RtfSpec.RtfVersion1;
			defaultFontId = "f0";
			fontTable.Clear();
			colorTable.Clear();
			generator = null;
			uniqueTextFormats.Clear();
			textFormatStack.Clear();
			currentTextFormat = null;
			documentInfo.Reset();
			userProperties.Clear();
		} // Reset

		// ----------------------------------------------------------------------
		// members
		private RtfInterpreterState state;
		private int rtfVersion;
		private string defaultFontId;
		private readonly RtfFontCollection fontTable = new RtfFontCollection();
		private readonly RtfColorCollection colorTable = new RtfColorCollection();
		private string generator;
		private readonly RtfTextFormatCollection uniqueTextFormats = new RtfTextFormatCollection();
		private readonly Stack textFormatStack = new Stack();
		private RtfTextFormat currentTextFormat;
		private readonly RtfDocumentInfo documentInfo = new RtfDocumentInfo();
		private readonly RtfDocumentPropertyCollection userProperties = new RtfDocumentPropertyCollection();

	} // class RtfInterpreterContext

} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------
