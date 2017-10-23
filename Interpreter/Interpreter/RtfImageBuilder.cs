// -- FILE ------------------------------------------------------------------
// name       : RtfImageBuilder.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using Itenso.Rtf.Support;

namespace Itenso.Rtf.Interpreter
{

	// ------------------------------------------------------------------------
	public sealed class RtfImageBuilder : RtfElementVisitorBase
	{

		// ----------------------------------------------------------------------
		public RtfImageBuilder() :
			base( RtfElementVisitorOrder.DepthFirst )
		{
			Reset();
		} // RtfImageBuilder

		// ----------------------------------------------------------------------
		public void Reset()
		{
			format = RtfVisualImageFormat.Bmp;
			width = 0;
			height = 0;
			desiredWidth = 0;
			desiredHeight = 0;
			scaleWidthPercent = 100;
			scaleHeightPercent = 100;
			imageDataHex = null;
		} // Reset

		// ----------------------------------------------------------------------
		public RtfVisualImageFormat Format
		{
			get { return format; }
		} // Format

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
		protected override void DoVisitGroup( IRtfGroup group )
		{
			switch ( group.Destination )
			{
				case RtfSpec.TagPicture:
					Reset();
					VisitGroupChildren( group );
					break;
			}
		} // DoVisitGroup

		// ----------------------------------------------------------------------
		protected override void DoVisitTag( IRtfTag tag )
		{
			switch ( tag.Name )
			{
				case RtfSpec.TagPictureFormatWinDib:
				case RtfSpec.TagPictureFormatWinBmp:
					format = RtfVisualImageFormat.Bmp;
					break;
				case RtfSpec.TagPictureFormatEmf:
					format = RtfVisualImageFormat.Emf;
					break;
				case RtfSpec.TagPictureFormatJpg:
					format = RtfVisualImageFormat.Jpg;
					break;
				case RtfSpec.TagPictureFormatPng:
					format = RtfVisualImageFormat.Png;
					break;
				case RtfSpec.TagPictureFormatWmf:
					format = RtfVisualImageFormat.Wmf;
					break;
				case RtfSpec.TagPictureWidth:
					width = Math.Abs( tag.ValueAsNumber );
					desiredWidth = width;
					break;
				case RtfSpec.TagPictureHeight:
					height = Math.Abs( tag.ValueAsNumber );
					desiredHeight = height;
					break;
				case RtfSpec.TagPictureWidthGoal:
					desiredWidth = Math.Abs( tag.ValueAsNumber );
					if ( width == 0 )
					{
						// hack to prevent WordPad documents which lack the \picw and \pich tags
						// from resulting in an exception due to undefined width/height
						width = desiredWidth;
					}
					break;
				case RtfSpec.TagPictureHeightGoal:
					desiredHeight = Math.Abs( tag.ValueAsNumber );
					if ( height == 0 )
					{
						// hack to prevent WordPad documents which lack the \picw and \pich tags
						// from resulting in an exception due to undefined width/height
						height = desiredHeight;
					}
					break;
				case RtfSpec.TagPictureWidthScale:
					scaleWidthPercent = Math.Abs( tag.ValueAsNumber );
					break;
				case RtfSpec.TagPictureHeightScale:
					scaleHeightPercent = Math.Abs( tag.ValueAsNumber );
					break;
			}
		} // DoVisitTag

		// ----------------------------------------------------------------------
		protected override void DoVisitText( IRtfText text )
		{
			imageDataHex = text.Text;
		} // DoVisitText

		// ----------------------------------------------------------------------
		// members
		private RtfVisualImageFormat format;
		private int width;
		private int height;
		private int desiredWidth;
		private int desiredHeight;
		private int scaleWidthPercent;
		private int scaleHeightPercent;
		private string imageDataHex;

	} // class RtfImageBuilder

} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------
