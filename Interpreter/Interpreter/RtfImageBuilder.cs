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
        // members

        // ----------------------------------------------------------------------
        public RtfVisualImageFormat Format { get; private set; } // Format

        // ----------------------------------------------------------------------
        public int Width { get; private set; } // Width

        // ----------------------------------------------------------------------
        public int Height { get; private set; } // Height

        // ----------------------------------------------------------------------
        public int DesiredWidth { get; private set; } // DesiredWidth

        // ----------------------------------------------------------------------
        public int DesiredHeight { get; private set; } // DesiredHeight

        // ----------------------------------------------------------------------
        public int ScaleWidthPercent { get; private set; } // ScaleWidthPercent

        // ----------------------------------------------------------------------
        public int ScaleHeightPercent { get; private set; } // ScaleHeightPercent

        // ----------------------------------------------------------------------
        public string ImageDataHex { get; private set; } // ImageDataHex

        // ----------------------------------------------------------------------
        public RtfImageBuilder() :
            base(RtfElementVisitorOrder.DepthFirst)
        {
            Reset();
        } // RtfImageBuilder

        // ----------------------------------------------------------------------
        public void Reset()
        {
            Format = RtfVisualImageFormat.Bmp;
            Width = 0;
            Height = 0;
            DesiredWidth = 0;
            DesiredHeight = 0;
            ScaleWidthPercent = 100;
            ScaleHeightPercent = 100;
            ImageDataHex = null;
        } // Reset

        // ----------------------------------------------------------------------
        protected override void DoVisitGroup(IRtfGroup group)
        {
            switch (group.Destination)
            {
                case RtfSpec.TagPicture:
                    Reset();
                    VisitGroupChildren(group);
                    break;
            }
        } // DoVisitGroup

        // ----------------------------------------------------------------------
        protected override void DoVisitTag(IRtfTag tag)
        {
            switch (tag.Name)
            {
                case RtfSpec.TagPictureFormatWinDib:
                case RtfSpec.TagPictureFormatWinBmp:
                    Format = RtfVisualImageFormat.Bmp;
                    break;
                case RtfSpec.TagPictureFormatEmf:
                    Format = RtfVisualImageFormat.Emf;
                    break;
                case RtfSpec.TagPictureFormatJpg:
                    Format = RtfVisualImageFormat.Jpg;
                    break;
                case RtfSpec.TagPictureFormatPng:
                    Format = RtfVisualImageFormat.Png;
                    break;
                case RtfSpec.TagPictureFormatWmf:
                    Format = RtfVisualImageFormat.Wmf;
                    break;
                case RtfSpec.TagPictureWidth:
                    Width = Math.Abs(tag.ValueAsNumber);
                    DesiredWidth = Width;
                    break;
                case RtfSpec.TagPictureHeight:
                    Height = Math.Abs(tag.ValueAsNumber);
                    DesiredHeight = Height;
                    break;
                case RtfSpec.TagPictureWidthGoal:
                    DesiredWidth = Math.Abs(tag.ValueAsNumber);
                    if (Width == 0)
                        Width = DesiredWidth;
                    break;
                case RtfSpec.TagPictureHeightGoal:
                    DesiredHeight = Math.Abs(tag.ValueAsNumber);
                    if (Height == 0)
                        Height = DesiredHeight;
                    break;
                case RtfSpec.TagPictureWidthScale:
                    ScaleWidthPercent = Math.Abs(tag.ValueAsNumber);
                    break;
                case RtfSpec.TagPictureHeightScale:
                    ScaleHeightPercent = Math.Abs(tag.ValueAsNumber);
                    break;
            }
        } // DoVisitTag

        // ----------------------------------------------------------------------
        protected override void DoVisitText(IRtfText text)
        {
            ImageDataHex = text.Text;
        } // DoVisitText
    } // class RtfImageBuilder
} // namespace Itenso.Rtf.Interpreter
// -- EOF -------------------------------------------------------------------