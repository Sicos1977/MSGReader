// name       : RtfInterpreterListenerFileLogger.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Globalization;
using System.IO;

namespace Itenso.Rtf.Interpreter
{
    public class RtfInterpreterListenerFileLogger : RtfInterpreterListenerBase, IDisposable
    {
        public const string DefaultLogFileExtension = ".interpreter.log";

        // members
        private StreamWriter streamWriter;

        public string FileName { get; } // FileName

        public RtfInterpreterLoggerSettings Settings { get; } // Settings

        public RtfInterpreterListenerFileLogger(string fileName) :
            this(fileName, new RtfInterpreterLoggerSettings())
        {
        } // RtfInterpreterListenerFileLogger

        public RtfInterpreterListenerFileLogger(string fileName, RtfInterpreterLoggerSettings settings)
        {
            if (fileName == null)
                throw new ArgumentNullException("fileName");
            if (settings == null)
                throw new ArgumentNullException("settings");

            FileName = fileName;
            Settings = settings;
        } // RtfInterpreterListenerFileLogger

        public virtual void Dispose()
        {
            CloseStream();
        } // Dispose

        protected override void DoBeginDocument(IRtfInterpreterContext context)
        {
            EnsureDirectory();
            OpenStream();

            if (Settings.Enabled && !string.IsNullOrEmpty(Settings.BeginDocumentText))
                WriteLine(Settings.BeginDocumentText);
        } // DoBeginDocument

        protected override void DoInsertText(IRtfInterpreterContext context, string text)
        {
            if (Settings.Enabled && !string.IsNullOrEmpty(Settings.TextFormatText))
            {
                var msg = text;
                if (msg.Length > Settings.TextMaxLength && !string.IsNullOrEmpty(Settings.TextOverflowText))
                    msg = msg.Substring(0, msg.Length - Settings.TextOverflowText.Length) + Settings.TextOverflowText;
                WriteLine(string.Format(
                    CultureInfo.InvariantCulture,
                    Settings.TextFormatText,
                    msg,
                    context.CurrentTextFormat));
            }
        } // DoInsertText

        protected override void DoInsertSpecialChar(IRtfInterpreterContext context, RtfVisualSpecialCharKind kind)
        {
            if (Settings.Enabled && !string.IsNullOrEmpty(Settings.SpecialCharFormatText))
                WriteLine(string.Format(
                    CultureInfo.InvariantCulture,
                    Settings.SpecialCharFormatText,
                    kind));
        } // DoInsertSpecialChar

        protected override void DoInsertBreak(IRtfInterpreterContext context, RtfVisualBreakKind kind)
        {
            if (Settings.Enabled && !string.IsNullOrEmpty(Settings.BreakFormatText))
                WriteLine(string.Format(
                    CultureInfo.InvariantCulture,
                    Settings.BreakFormatText,
                    kind));
        } // DoInsertBreak

        protected override void DoInsertImage(IRtfInterpreterContext context,
            RtfVisualImageFormat format,
            int width, int height, int desiredWidth, int desiredHeight,
            int scaleWidthPercent, int scaleHeightPercent,
            string imageDataHex
        )
        {
            if (Settings.Enabled && !string.IsNullOrEmpty(Settings.ImageFormatText))
                WriteLine(string.Format(
                    CultureInfo.InvariantCulture,
                    Settings.ImageFormatText,
                    format,
                    width,
                    height,
                    desiredWidth,
                    desiredHeight,
                    scaleWidthPercent,
                    scaleHeightPercent,
                    imageDataHex,
                    imageDataHex.Length / 2));
        } // DoInsertImage

        protected override void DoEndDocument(IRtfInterpreterContext context)
        {
            if (Settings.Enabled && !string.IsNullOrEmpty(Settings.EndDocumentText))
                WriteLine(Settings.EndDocumentText);

            CloseStream();
        } // DoEndDocument

        private void WriteLine(string message)
        {
            if (streamWriter == null)
                return;

            streamWriter.WriteLine(message);
            streamWriter.Flush();
        } // WriteLine

        private void EnsureDirectory()
        {
            var fi = new FileInfo(FileName);
            if (!string.IsNullOrEmpty(fi.DirectoryName) && !Directory.Exists(fi.DirectoryName))
                Directory.CreateDirectory(fi.DirectoryName);
        } // EnsureDirectory

        private void OpenStream()
        {
            if (streamWriter != null)
                return;
            streamWriter = new StreamWriter(FileName);
        } // OpenStream

        private void CloseStream()
        {
            if (streamWriter == null)
                return;
            streamWriter.Close();
            streamWriter.Dispose();
            streamWriter = null;
        } // OpenStream
    } // class RtfInterpreterListenerFileLogger
}