// name       : RtfInterpreterBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Collections;

namespace Itenso.Rtf.Interpreter
{
    public abstract class RtfInterpreterBase : IRtfInterpreter
    {
        // Members
        private ArrayList _listeners;

        protected RtfInterpreterContext Context { get; } = new RtfInterpreterContext();

// Context

        protected RtfInterpreterBase(params IRtfInterpreterListener[] listeners) :
            this(new RtfInterpreterSettings(), listeners)
        {
        } // RtfInterpreterBase

        protected RtfInterpreterBase(IRtfInterpreterSettings settings, params IRtfInterpreterListener[] listeners)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            Settings = settings;
            if (listeners != null)
                foreach (var listener in listeners)
                    AddInterpreterListener(listener);
        } // RtfInterpreterBase

        public IRtfInterpreterSettings Settings { get; } // Settings

        public void AddInterpreterListener(IRtfInterpreterListener listener)
        {
            if (listener == null)
                throw new ArgumentNullException(nameof(listener));
            if (_listeners == null)
                _listeners = new ArrayList();
            if (!_listeners.Contains(listener))
                _listeners.Add(listener);
        } // AddInterpreterListener

        public void RemoveInterpreterListener(IRtfInterpreterListener listener)
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
        } // RemoveInterpreterListener

        public void Interpret(IRtfGroup rtfDocument)
        {
            if (rtfDocument == null)
                throw new ArgumentNullException(nameof(rtfDocument));
            DoInterpret(rtfDocument);
        } // Interpret

        protected abstract void DoInterpret(IRtfGroup rtfDocument);

        protected void NotifyBeginDocument()
        {
            if (_listeners != null)
                foreach (IRtfInterpreterListener listener in _listeners)
                    listener.BeginDocument(Context);
        } // NotifyBeginDocument

        protected void NotifyInsertText(string text)
        {
            if (_listeners != null)
                foreach (IRtfInterpreterListener listener in _listeners)
                    listener.InsertText(Context, text);
        } // NotifyInsertText

        protected void NotifyInsertSpecialChar(RtfVisualSpecialCharKind kind)
        {
            if (_listeners != null)
                foreach (IRtfInterpreterListener listener in _listeners)
                    listener.InsertSpecialChar(Context, kind);
        } // NotifyInsertSpecialChar

        protected void NotifyInsertBreak(RtfVisualBreakKind kind)
        {
            if (_listeners != null)
                foreach (IRtfInterpreterListener listener in _listeners)
                    listener.InsertBreak(Context, kind);
        } // NotifyInsertBreak

        protected void NotifyInsertImage(RtfVisualImageFormat format,
            int width, int height, int desiredWidth, int desiredHeight,
            int scaleWidthPercent, int scaleHeightPercent, string imageDataHex
        )
        {
            if (_listeners != null)
                foreach (IRtfInterpreterListener listener in _listeners)
                    listener.InsertImage(
                        Context,
                        format,
                        width,
                        height,
                        desiredWidth,
                        desiredHeight,
                        scaleWidthPercent,
                        scaleHeightPercent,
                        imageDataHex);
        } // NotifyInsertImage

        protected void NotifyEndDocument()
        {
            if (_listeners != null)
                foreach (IRtfInterpreterListener listener in _listeners)
                    listener.EndDocument(Context);
        } // NotifyEndDocument
    } // class RtfInterpreterBase
}