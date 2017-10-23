// -- FILE ------------------------------------------------------------------
// name       : RtfVisualImage.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Globalization;
using System.IO;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfVisualImage : RtfVisual, IRtfVisualImage
	{

		// ----------------------------------------------------------------------
		public RtfVisualImage(
			RtfVisualImageFormat format,
			RtfTextAlignment alignment,
			int width,
			int height,
			int desiredWidth,
			int desiredHeight,
			int scaleWidthPercent,
			int scaleHeightPercent,
			string imageDataHex
		) :
			base( RtfVisualKind.Image )
		{
			if ( width <= 0 )
			{
				throw new ArgumentException( Strings.InvalidImageWidth( width ) );
			}
			if ( height <= 0 )
			{
				throw new ArgumentException( Strings.InvalidImageHeight( height ) );
			}
			if ( desiredWidth <= 0 )
			{
				throw new ArgumentException( Strings.InvalidImageDesiredWidth( desiredWidth ) );
			}
			if ( desiredHeight <= 0 )
			{
				throw new ArgumentException( Strings.InvalidImageDesiredHeight( desiredHeight ) );
			}
			if ( scaleWidthPercent <= 0 )
			{
				throw new ArgumentException( Strings.InvalidImageScaleWidth( scaleWidthPercent ) );
			}
			if ( scaleHeightPercent <= 0 )
			{
				throw new ArgumentException( Strings.InvalidImageScaleHeight( scaleHeightPercent ) );
			}
			if ( imageDataHex == null )
			{
				throw new ArgumentNullException( "imageDataHex" );
			}
			this.format = format;
			this.alignment = alignment;
			this.width = width;
			this.height = height;
			this.desiredWidth = desiredWidth;
			this.desiredHeight = desiredHeight;
			this.scaleWidthPercent = scaleWidthPercent;
			this.scaleHeightPercent = scaleHeightPercent;
			this.imageDataHex = imageDataHex;
		} // RtfVisualImage

		// ----------------------------------------------------------------------
		protected override void DoVisit( IRtfVisualVisitor visitor )
		{
			visitor.VisitImage( this );
		} // DoVisit

		// ----------------------------------------------------------------------
		public RtfVisualImageFormat Format
		{
			get { return format; }
		} // Format

		// ----------------------------------------------------------------------
		public RtfTextAlignment Alignment
		{
			get { return alignment; }
			set { alignment = value; }
		} // Alignment

		// ----------------------------------------------------------------------
		public int Width
		{
			get { return width; }
		} // Width

		// ----------------------------------------------------------------------
		public int Height
		{
			get { return height; }
		} // Height

		// ----------------------------------------------------------------------
		public int DesiredWidth
		{
			get { return desiredWidth; }
		} // DesiredWidth

		// ----------------------------------------------------------------------
		public int DesiredHeight
		{
			get { return desiredHeight; }
		} // DesiredHeight

		// ----------------------------------------------------------------------
		public int ScaleWidthPercent
		{
			get { return scaleWidthPercent; }
		} // ScaleWidthPercent

		// ----------------------------------------------------------------------
		public int ScaleHeightPercent
		{
			get { return scaleHeightPercent; }
		} // ScaleHeightPercent

		// ----------------------------------------------------------------------
		public string ImageDataHex
		{
			get { return imageDataHex; }
		} // ImageDataHex

		// ----------------------------------------------------------------------
		public byte[] ImageDataBinary
		{
			get { return imageDataBinary ?? ( imageDataBinary = ToBinary( imageDataHex ) ); }
		} // ImageDataBinary

		// ----------------------------------------------------------------------
		public System.Drawing.Image ImageForDrawing
		{
			get
			{
				switch ( format )
				{
					case RtfVisualImageFormat.Bmp:
					case RtfVisualImageFormat.Jpg:
					case RtfVisualImageFormat.Png:
					case RtfVisualImageFormat.Emf:
					case RtfVisualImageFormat.Wmf:
						byte[] data = ImageDataBinary;
						return System.Drawing.Image.FromStream( new MemoryStream( data, 0, data.Length ) );
				}
				return null;
			}
		} // ImageForDrawing

		// ----------------------------------------------------------------------
		public static byte[] ToBinary( string imageDataHex )
		{
			if ( imageDataHex == null )
			{
				throw new ArgumentNullException( "imageDataHex" );
			}

			int hexDigits = imageDataHex.Length;
			int dataSize = hexDigits / 2;
			byte[] imageDataBinary = new byte[ dataSize ];

			StringBuilder hex = new StringBuilder( 2 );

			int dataPos = 0;
			for ( int i = 0; i < hexDigits; i++ )
			{
				char c = imageDataHex[ i ];
				if ( char.IsWhiteSpace( c ) )
				{
					continue;
				}
				hex.Append( imageDataHex[ i ] );
				if ( hex.Length == 2 )
				{
					imageDataBinary[ dataPos ] = byte.Parse( hex.ToString(), NumberStyles.HexNumber );
					dataPos++;
					hex.Remove( 0, 2 );
				}
			}

			return imageDataBinary;
		} // ToBinary

		// ----------------------------------------------------------------------
		protected override bool IsEqual( object obj )
		{
			RtfVisualImage compare = obj as RtfVisualImage; // guaranteed to be non-null
			return
				compare != null &&
				base.IsEqual( compare ) &&
				format == compare.format &&
				alignment == compare.alignment &&
				width == compare.width &&
				height == compare.height &&
				desiredWidth == compare.desiredWidth &&
				desiredHeight == compare.desiredHeight &&
				scaleWidthPercent == compare.scaleWidthPercent &&
				scaleHeightPercent == compare.scaleHeightPercent &&
				imageDataHex.Equals( compare.imageDataHex );
			//imageDataBinary.Equals( compare.imageDataBinary ); // cached info only
		} // IsEqual

		// ----------------------------------------------------------------------
		protected override int ComputeHashCode()
		{
			int hash = base.ComputeHashCode();
			hash = HashTool.AddHashCode( hash, format );
			hash = HashTool.AddHashCode( hash, alignment );
			hash = HashTool.AddHashCode( hash, width );
			hash = HashTool.AddHashCode( hash, height );
			hash = HashTool.AddHashCode( hash, desiredWidth );
			hash = HashTool.AddHashCode( hash, desiredHeight );
			hash = HashTool.AddHashCode( hash, scaleWidthPercent );
			hash = HashTool.AddHashCode( hash, scaleHeightPercent );
			hash = HashTool.AddHashCode( hash, imageDataHex );
			//hash = HashTool.AddHashCode( hash, imageDataBinary ); // cached info only
			return hash;
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return "[" + format + ": " + alignment + ", " +
				width + " x " + height + " " +
				"(" + desiredWidth + " x " + desiredHeight + ") " +
				"{" + scaleWidthPercent + "% x " + scaleHeightPercent + "%} " +
				":" + ( imageDataHex.Length / 2 ) + " bytes]";
		} // ToString

		// ----------------------------------------------------------------------
		// members
		private readonly RtfVisualImageFormat format;
		private RtfTextAlignment alignment;
		private readonly int width;
		private readonly int height;
		private readonly int desiredWidth;
		private readonly int desiredHeight;
		private readonly int scaleWidthPercent;
		private readonly int scaleHeightPercent;
		private readonly string imageDataHex;
		private byte[] imageDataBinary; // cached info only

	} // class RtfVisualImage

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
