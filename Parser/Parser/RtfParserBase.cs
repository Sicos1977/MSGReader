// name       : RtfParserBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;

namespace Itenso.Rtf.Parser
{
    public abstract class RtfParserBase : IRtfParser
    {
        // Members
        private ArrayList _listeners;

        protected RtfParserBase()
        {
        } // RtfParserBase

        protected RtfParserBase(params IRtfParserListener[] listeners)
        {
            if (listeners != null)
                foreach (var listener in listeners)
                    AddParserListener(listener);
        } // RtfParserBase

        public bool IgnoreContentAfterRootGroup { get; set; }

        public void AddParserListener(IRtfParserListener listener)
        {
            if (listener == null)
                throw new ArgumentNullException(nameof(listener));
            if (_listeners == null)
                _listeners = new ArrayList();
            if (!_listeners.Contains(listener))
                _listeners.Add(listener);
        } // AddParserListener

        public void RemoveParserListener(IRtfParserListener listener)
        {
            if (listener == null)
                throw new ArgumentNullException(nameof(listener));
            if (_listeners != null)
            {
                if (_listeners.Contains(listener))
                    _listeners.Remove(listener);
                if (_listeners.Count == 0)
                    _listeners = null;
            }
        } // RemoveParserListener

        public void Parse(IRtfSource rtfTextSource)
        {
            if (rtfTextSource == null)
                throw new ArgumentNullException(nameof(rtfTextSource));
            DoParse(rtfTextSource);
        } // Parse

        protected abstract void DoParse(IRtfSource rtfTextSource);

        protected void NotifyParseBegin()
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.ParseBegin();
        } // NotifyParseBegin

        protected void NotifyGroupBegin()
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.GroupBegin();
        } // NotifyGroupBegin

        protected void NotifyTagFound(IRtfTag tag)
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.TagFound(tag);
        } // NotifyTagFound

        protected void NotifyTextFound(IRtfText text)
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.TextFound(text);
        } // NotifyTextFound

        protected void NotifyGroupEnd()
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.GroupEnd();
        } // NotifyGroupEnd

        protected void NotifyParseSuccess()
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.ParseSuccess();
        } // NotifyParseSuccess

        protected void NotifyParseFail(RtfException reason)
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.ParseFail(reason);
        } // NotifyParseFail

        protected void NotifyParseEnd()
        {
            if (_listeners != null)
                foreach (IRtfParserListener listener in _listeners)
                    listener.ParseEnd();
        } // NotifyParseEnd
    } // class RtfParserBase
}