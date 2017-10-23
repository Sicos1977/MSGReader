// -- FILE ------------------------------------------------------------------
// name       : RtfInterpreterListenerFileLogger.cs
// project    : RTF Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.IO;
using System.Globalization;

namespace Itenso.Rtf.Interpreter
{

	// ------------------------------------------------------------------------
	public class RtfInterpreterListenerFileLogger : RtfInterpreterListenerBase, IDisposable
	{

		// ----------------------------------------------------------------------
		public const string DefaultLogFileExtension = ".interpreter.log";

		// ----------------------------------------------------------------------
		public RtfInterpreterListenerFileLogger( string fileName ) :
			this( fileName, new RtfInterpreterLoggerSettings() )
		{
		} // RtfInterpreterListenerFileLogger

		// ----------------------------------------------------------------------
		public RtfInterpreterListenerFileLogger( string fileName, RtfInterpreterLoggerSettings settings )
		{
			if ( fileName == null )
			{
				throw new ArgumentNullException( "fileName" );
			}
			if ( settings == null )
			{
				throw new ArgumentNullException( "settings" );
			}

			this.fileName = fileName;
			this.settings = settings;
		} // RtfInterpreterListenerFileLogger

		// ----------------------------------------------------------------------
		public string FileName
		{
			get { return fileName; }
		} // FileName

		// ----------------------------------------------------------------------
		public RtfInterpreterLoggerSettings Settings
		{
			get { return settings; }
		} // Settings

		// ----------------------------------------------------------------------
		public virtual void Dispose()
		{
			CloseStream();
		} // Dispose

		// ----------------------------------------------------------------------
		protected override void DoBeginDocument( IRtfInterpreterContext context )
		{
			EnsureDirectory();
			OpenStream();

			if ( settings.Enabled && !string.IsNullOrEmpty( settings.BeginDocumentText ) )
			{
				WriteLine( settings.BeginDocumentText );
			}
		} // DoBeginDocument

		// ----------------------------------------------------------------------
		protected override void DoInsertText( IRtfInterpreterContext context, string text )
		{
			if ( settings.Enabled && !string.IsNullOrEmpty( settings.TextFormatText ) )
			{
				string msg = text;
				if ( msg.Length > settings.TextMaxLength && !string.IsNullOrEmpty( settings.TextOverflowText ) )
				{
					msg = msg.Substring( 0, msg.Length - settings.TextOverflowText.Length ) + settings.TextOverflowText;
				}
				WriteLine( string.Format(
					CultureInfo.InvariantCulture,
					settings.TextFormatText,
					msg,
					context.CurrentTextFormat ) );
			}
		} // DoInsertText

		// ----------------------------------------------------------------------
		protected override void DoInsertSpecialChar( IRtfInterpreterContext context, RtfVisualSpecialCharKind kind )
		{
			if ( settings.Enabled && !string.IsNullOrEmpty( settings.SpecialCharFormatText ) )
			{
				WriteLine( string.Format(
					CultureInfo.InvariantCulture,
					settings.SpecialCharFormatText,
					kind ) );
			}
		} // DoInsertSpecialChar

		// ----------------------------------------------------------------------
		protected override void DoInsertBreak( IRtfInterpreterContext context, RtfVisualBreakKind kind )
		{
			if ( settings.Enabled && !string.IsNullOrEmpty( settings.BreakFormatText ) )
			{
				WriteLine( string.Format(
					CultureInfo.InvariantCulture,
					settings.BreakFormatText,
					kind ) );
			}
		} // DoInsertBreak

		// ----------------------------------------------------------------------
		protected override void DoInsertImage( IRtfInterpreterContext context,
			RtfVisualImageFormat format,
			int width, int height, int desiredWidth, int desiredHeight,
			int scaleWidthPercent, int scaleHeightPercent,
			string imageDataHex
		)
		{
			if ( settings.Enabled && !string.IsNullOrEmpty( settings.ImageFormatText ) )
			{
				WriteLine( string.Format(
					CultureInfo.InvariantCulture,
					settings.ImageFormatText,
					format,
					width,
					height,
					desiredWidth,
					desiredHeight,
					scaleWidthPercent,
					scaleHeightPercent,
					imageDataHex,
					(imageDataHex.Length / 2) ) );
			}
		} // DoInsertImage

		// ----------------------------------------------------------------------
		protected override void DoEndDocument( IRtfInterpreterContext context )
		{
			if ( settings.Enabled && !string.IsNullOrEmpty( settings.EndDocumentText ) )
			{
				WriteLine( settings.EndDocumentText );
			}

			CloseStream();
		} // DoEndDocument

		// ----------------------------------------------------------------------
		private void WriteLine( string message )
		{
			if ( streamWriter == null )
			{
				return;
			}

			streamWriter.WriteLine( message );
			streamWriter.Flush();
		} // WriteLine

		// ----------------------------------------------------------------------
		private void EnsureDirectory()
		{
			FileInfo fi = new FileInfo( fileName );
			if ( !string.IsNullOrEmpty( fi.DirectoryName ) && !Directory.Exists( fi.DirectoryName ) )
			{
				Directory.CreateDirectory( fi.DirectoryName );
			}
		} // EnsureDirectory

		// ----------------------------------------------------------------------
		private void OpenStream()
		{
			if ( streamWriter != null )
			{
				return;
			}
			streamWriter = new StreamWriter( fileName );
		} // OpenStream

		// ----------------------------------------------------------------------
		private void CloseStream()
		{
			if ( streamWriter == null )
			{
				return;
			}
			streamWriter.Close();
			streamWriter.Dispose();
			streamWriter = null;
		} // OpenStream

		// ----------------------------------------------------------------------
		// members
		private readonly string fileName;
		private readonly RtfInterpreterLoggerSettings settings;
		private StreamWriter streamWriter;

	} // class RtfInterpreterListenerFileLogger

} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------
	