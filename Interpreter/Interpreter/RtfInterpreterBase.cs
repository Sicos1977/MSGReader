// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreterBase.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System;
using System.Collections;

namespace Itenso.Rtf.Interpreter
{
    // ------------------------------------------------------------------------
    public abstract class RtfInterpreterBase : IRtfInterpreter
    {
        // ----------------------------------------------------------------------
        // members
        private ArrayList listeners;

        // ----------------------------------------------------------------------
        protected RtfInterpreterContext Context { get; } = new RtfInterpreterContext();

// Context

        // ----------------------------------------------------------------------
        protected RtfInterpreterBase(params IRtfInterpreterListener[] listeners) :
            this(new RtfInterpreterSettings(), listeners)
        {
        } // RtfInterpreterBase

        // ----------------------------------------------------------------------
        protected RtfInterpreterBase(IRtfInterpreterSettings settings, params IRtfInterpreterListener[] listeners)
        {
            if (settings == null)
                throw new ArgumentNullException("settings");

            Settings = settings;
            if (listeners != null)
                foreach (var listener in listeners)
                    AddInterpreterListener(listener);
        } // RtfInterpreterBase

        // ----------------------------------------------------------------------
        public IRtfInterpreterSettings Settings { get; } // Settings

        // ----------------------------------------------------------------------
        public void AddInterpreterListener(IRtfInterpreterListener listener)
        {
            if (listener == null)
                throw new ArgumentNullException("listener");
            if (listeners == null)
                listeners = new ArrayList();
            if (!listeners.Contains(listener))
                listeners.Add(listener);
        } // AddInterpreterListener

        // ----------------------------------------------------------------------
        public void RemoveInterpreterListener(IRtfInterpreterListener listener)
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
        } // RemoveInterpreterListener

        // ----------------------------------------------------------------------
        public void Interpret(IRtfGroup rtfDocument)
        {
            if (rtfDocument == null)
                throw new ArgumentNullException("rtfDocument");
            DoInterpret(rtfDocument);
        } // Interpret

        // ----------------------------------------------------------------------
        protected abstract void DoInterpret(IRtfGroup rtfDocument);

        // ----------------------------------------------------------------------
        protected void NotifyBeginDocument()
        {
            if (listeners != null)
                foreach (IRtfInterpreterListener listener in listeners)
                    listener.BeginDocument(Context);
        } // NotifyBeginDocument

        // ----------------------------------------------------------------------
        protected void NotifyInsertText(string text)
        {
            if (listeners != null)
                foreach (IRtfInterpreterListener listener in listeners)
                    listener.InsertText(Context, text);
        } // NotifyInsertText

        // ----------------------------------------------------------------------
        protected void NotifyInsertSpecialChar(RtfVisualSpecialCharKind kind)
        {
            if (listeners != null)
                foreach (IRtfInterpreterListener listener in listeners)
                    listener.InsertSpecialChar(Context, kind);
        } // NotifyInsertSpecialChar

        // ----------------------------------------------------------------------
        protected void NotifyInsertBreak(RtfVisualBreakKind kind)
        {
            if (listeners != null)
                foreach (IRtfInterpreterListener listener in listeners)
                    listener.InsertBreak(Context, kind);
        } // NotifyInsertBreak

        // ----------------------------------------------------------------------
        protected void NotifyInsertImage(RtfVisualImageFormat format,
            int width, int height, int desiredWidth, int desiredHeight,
            int scaleWidthPercent, int scaleHeightPercent, string imageDataHex
        )
        {
            if (listeners != null)
                foreach (IRtfInterpreterListener listener in listeners)
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

        // ----------------------------------------------------------------------
        protected void NotifyEndDocument()
        {
            if (listeners != null)
                foreach (IRtfInterpreterListener listener in listeners)
                    listener.EndDocument(Context);
        } // NotifyEndDocument
    } // class RtfInterpreterBase
} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------