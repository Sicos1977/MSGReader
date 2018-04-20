// -- FILE ------------------------------------------------------------------
// name       : RtfParserBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using System.Collections;

namespace Itenso.Rtf.Parser
{
    // ------------------------------------------------------------------------
    public abstract class RtfParserBase : IRtfParser
    {
        // ----------------------------------------------------------------------
        // members
        private ArrayList listeners;

        // ----------------------------------------------------------------------
        protected RtfParserBase()
        {
        } // RtfParserBase

        // ----------------------------------------------------------------------
        protected RtfParserBase(params IRtfParserListener[] listeners)
        {
            if (listeners != null)
                foreach (var listener in listeners)
                    AddParserListener(listener);
        } // RtfParserBase

        // ----------------------------------------------------------------------
        public bool IgnoreContentAfterRootGroup { get; set; }

        // ----------------------------------------------------------------------
        public void AddParserListener(IRtfParserListener listener)
        {
            if (listener == null)
                throw new ArgumentNullException("listener");
            if (listeners == null)
                listeners = new ArrayList();
            if (!listeners.Contains(listener))
                listeners.Add(listener);
        } // AddParserListener

        // ----------------------------------------------------------------------
        public void RemoveParserListener(IRtfParserListener listener)
        {
            if (listener == null)
                throw new ArgumentNullException("listener");
            if (listeners != null)
            {
                if (listeners.Contains(listener))
                    listeners.Remove(listener);
                if (listeners.Count == 0)
                    listeners = null;
            }
        } // RemoveParserListener

        // ----------------------------------------------------------------------
        public void Parse(IRtfSource rtfTextSource)
        {
            if (rtfTextSource == null)
                throw new ArgumentNullException("rtfTextSource");
            DoParse(rtfTextSource);
        } // Parse

        // ----------------------------------------------------------------------
        protected abstract void DoParse(IRtfSource rtfTextSource);

        // ----------------------------------------------------------------------
        protected void NotifyParseBegin()
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.ParseBegin();
        } // NotifyParseBegin

        // ----------------------------------------------------------------------
        protected void NotifyGroupBegin()
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.GroupBegin();
        } // NotifyGroupBegin

        // ----------------------------------------------------------------------
        protected void NotifyTagFound(IRtfTag tag)
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.TagFound(tag);
        } // NotifyTagFound

        // ----------------------------------------------------------------------
        protected void NotifyTextFound(IRtfText text)
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.TextFound(text);
        } // NotifyTextFound

        // ----------------------------------------------------------------------
        protected void NotifyGroupEnd()
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.GroupEnd();
        } // NotifyGroupEnd

        // ----------------------------------------------------------------------
        protected void NotifyParseSuccess()
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.ParseSuccess();
        } // NotifyParseSuccess

        // ----------------------------------------------------------------------
        protected void NotifyParseFail(RtfException reason)
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.ParseFail(reason);
        } // NotifyParseFail

        // ----------------------------------------------------------------------
        protected void NotifyParseEnd()
        {
            if (listeners != null)
                foreach (IRtfParserListener listener in listeners)
                    listener.ParseEnd();
        } // NotifyParseEnd
    } // class RtfParserBase
} // namespace Itenso.Rtf.Parser
// -- EOF -------------------------------------------------------------------