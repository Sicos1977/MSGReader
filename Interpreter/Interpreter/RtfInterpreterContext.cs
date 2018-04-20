// name       : RtfInterpreterContext.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;
using Itenso.Rtf.Model;

namespace Itenso.Rtf.Interpreter
{
    public sealed class RtfInterpreterContext : IRtfInterpreterContext
    {
        private readonly Stack textFormatStack = new Stack();
        private readonly RtfTextFormatCollection uniqueTextFormats = new RtfTextFormatCollection();
        private RtfTextFormat currentTextFormat;

        // members

        public RtfFontCollection WritableFontTable { get; } = new RtfFontCollection();

// WritableFontTable

        public RtfColorCollection WritableColorTable { get; } = new RtfColorCollection();

// WritableColorTable

        public RtfTextFormat WritableCurrentTextFormat
        {
            get
            {
                if (currentTextFormat == null)
                    WritableCurrentTextFormat = new RtfTextFormat(DefaultFont, RtfSpec.DefaultFontSize);
                return currentTextFormat;
            }
            set { currentTextFormat = (RtfTextFormat) GetUniqueTextFormatInstance(value); }
        } // WritableCurrentTextFormat

        public RtfDocumentInfo WritableDocumentInfo { get; } = new RtfDocumentInfo();

// WritableDocumentInfo

        public RtfDocumentPropertyCollection WritableUserProperties { get; } = new RtfDocumentPropertyCollection();

// WritableUserProperties

        public RtfInterpreterState State { get; set; } // State

        public int RtfVersion { get; set; } // RtfVersion

        public string DefaultFontId { get; set; } // DefaultFontIndex

        public IRtfFont DefaultFont
        {
            get
            {
                var defaultFont = WritableFontTable[DefaultFontId];
                if (defaultFont != null)
                    return defaultFont;
                throw new RtfUndefinedFontException(Strings.InvalidDefaultFont(
                    DefaultFontId, WritableFontTable.ToString()));
            }
        } // DefaultFont

        public IRtfFontCollection FontTable
        {
            get { return WritableFontTable; }
        } // FontTable

        public IRtfColorCollection ColorTable
        {
            get { return WritableColorTable; }
        } // ColorTable

        public string Generator { get; set; } // Generator

        public IRtfTextFormatCollection UniqueTextFormats
        {
            get { return uniqueTextFormats; }
        } // UniqueTextFormats

        public IRtfTextFormat CurrentTextFormat
        {
            get { return currentTextFormat; }
        } // CurrentTextFormat

        public IRtfTextFormat GetSafeCurrentTextFormat()
        {
            return currentTextFormat ?? WritableCurrentTextFormat;
        } // GetSafeCurrentTextFormat

        public IRtfTextFormat GetUniqueTextFormatInstance(IRtfTextFormat templateFormat)
        {
            if (templateFormat == null)
                throw new ArgumentNullException("templateFormat");
            IRtfTextFormat uniqueInstance;
            var existingEquivalentPos = uniqueTextFormats.IndexOf(templateFormat);
            if (existingEquivalentPos >= 0)
            {
                // we already know an equivalent format -> reference that one for future use
                uniqueInstance = uniqueTextFormats[existingEquivalentPos];
            }
            else
            {
                // this is a yet unknown format -> add it to the known formats and use it
                uniqueTextFormats.Add(templateFormat);
                uniqueInstance = templateFormat;
            }
            return uniqueInstance;
        } // GetUniqueTextFormatInstance

        public IRtfDocumentInfo DocumentInfo
        {
            get { return WritableDocumentInfo; }
        } // DocumentInfo

        public IRtfDocumentPropertyCollection UserProperties
        {
            get { return WritableUserProperties; }
        } // UserProperties

        public void PushCurrentTextFormat()
        {
            textFormatStack.Push(WritableCurrentTextFormat);
        } // PushCurrentTextFormat

        public void PopCurrentTextFormat()
        {
            if (textFormatStack.Count == 0)
                throw new RtfStructureException(Strings.InvalidTextContextState);
            currentTextFormat = (RtfTextFormat) textFormatStack.Pop();
        } // PopCurrentTextFormat

        public void Reset()
        {
            State = RtfInterpreterState.Init;
            RtfVersion = RtfSpec.RtfVersion1;
            DefaultFontId = "f0";
            WritableFontTable.Clear();
            WritableColorTable.Clear();
            Generator = null;
            uniqueTextFormats.Clear();
            textFormatStack.Clear();
            currentTextFormat = null;
            WritableDocumentInfo.Reset();
            WritableUserProperties.Clear();
        } // Reset
    } // class RtfInterpreterContext
}