// name       : RtfInterpreterTool.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.IO;
using Itenso.Rtf.Interpreter;

namespace Itenso.Rtf.Support
{
    public static class RtfInterpreterTool
    {
        public static IRtfDocument BuildDoc(string rtfText, params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(rtfText, new RtfInterpreterSettings(), listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(string rtfText, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(RtfParserTool.Parse(rtfText), settings, listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(TextReader rtfTextSource, params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(rtfTextSource, new RtfInterpreterSettings(), listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(TextReader rtfTextSource, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(RtfParserTool.Parse(rtfTextSource), settings, listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(Stream rtfTextSource, params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(rtfTextSource, new RtfInterpreterSettings(), listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(Stream rtfTextSource, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(RtfParserTool.Parse(rtfTextSource), settings, listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(IRtfSource rtfTextSource, params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(rtfTextSource, new RtfInterpreterSettings(), listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(IRtfSource rtfTextSource, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(RtfParserTool.Parse(rtfTextSource), settings, listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(IRtfGroup rtfDocument, params IRtfInterpreterListener[] listeners)
        {
            return BuildDoc(rtfDocument, new RtfInterpreterSettings(), listeners);
        } // BuildDoc

        public static IRtfDocument BuildDoc(IRtfGroup rtfDocument, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            var docBuilder = new RtfInterpreterListenerDocumentBuilder();
            IRtfInterpreterListener[] allListeners;
            if (listeners == null)
            {
                allListeners = new IRtfInterpreterListener[] {docBuilder};
            }
            else
            {
                allListeners = new IRtfInterpreterListener[listeners.Length + 1];
                allListeners[0] = docBuilder;
                listeners.CopyTo(allListeners, 1);
            }
            Interpret(rtfDocument, settings, allListeners);
            return docBuilder.Document;
        } // BuildDoc

        public static void Interpret(string rtfText, params IRtfInterpreterListener[] listeners)
        {
            Interpret(rtfText, new RtfInterpreterSettings(), listeners);
        } // Interpret

        public static void Interpret(string rtfText, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            Interpret(RtfParserTool.Parse(rtfText), settings, listeners);
        } // Interpret

        public static void Interpret(TextReader rtfTextSource, params IRtfInterpreterListener[] listeners)
        {
            Interpret(rtfTextSource, new RtfInterpreterSettings(), listeners);
        } // Interpret

        public static void Interpret(TextReader rtfTextSource, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            Interpret(RtfParserTool.Parse(rtfTextSource), settings, listeners);
        } // Interpret

        public static void Interpret(Stream rtfTextSource, params IRtfInterpreterListener[] listeners)
        {
            Interpret(rtfTextSource, new RtfInterpreterSettings(), listeners);
        } // Interpret

        public static void Interpret(Stream rtfTextSource, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            Interpret(RtfParserTool.Parse(rtfTextSource), settings, listeners);
        } // Interpret

        public static void Interpret(IRtfSource rtfTextSource, params IRtfInterpreterListener[] listeners)
        {
            Interpret(rtfTextSource, new RtfInterpreterSettings(), listeners);
        } // Interpret

        public static void Interpret(IRtfSource rtfTextSource, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            Interpret(RtfParserTool.Parse(rtfTextSource), settings, listeners);
        } // Interpret

        public static void Interpret(IRtfGroup rtfDocument, params IRtfInterpreterListener[] listeners)
        {
            Interpret(rtfDocument, new RtfInterpreterSettings(), listeners);
        } // Interpret

        public static void Interpret(IRtfGroup rtfDocument, IRtfInterpreterSettings settings,
            params IRtfInterpreterListener[] listeners)
        {
            var interpreter = new RtfInterpreter(settings);
            if (listeners != null)
                foreach (var listener in listeners)
                    if (listener != null)
                        interpreter.AddInterpreterListener(listener);
            interpreter.Interpret(rtfDocument);
        } // Interpret
    } // class RtfInterpreterTool
}