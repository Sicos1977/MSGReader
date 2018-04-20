// name       : RtfVisualImage.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.23
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{
    public sealed class RtfVisualImage : RtfVisual, IRtfVisualImage
    {
        // members
        private byte[] imageDataBinary; // cached info only

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
            base(RtfVisualKind.Image)
        {
            if (width <= 0)
                throw new ArgumentException(Strings.InvalidImageWidth(width));
            if (height <= 0)
                throw new ArgumentException(Strings.InvalidImageHeight(height));
            if (desiredWidth <= 0)
                throw new ArgumentException(Strings.InvalidImageDesiredWidth(desiredWidth));
            if (desiredHeight <= 0)
                throw new ArgumentException(Strings.InvalidImageDesiredHeight(desiredHeight));
            if (scaleWidthPercent <= 0)
                throw new ArgumentException(Strings.InvalidImageScaleWidth(scaleWidthPercent));
            if (scaleHeightPercent <= 0)
                throw new ArgumentException(Strings.InvalidImageScaleHeight(scaleHeightPercent));
            if (imageDataHex == null)
                throw new ArgumentNullException("imageDataHex");
            Format = format;
            Alignment = alignment;
            Width = width;
            Height = height;
            DesiredWidth = desiredWidth;
            DesiredHeight = desiredHeight;
            ScaleWidthPercent = scaleWidthPercent;
            ScaleHeightPercent = scaleHeightPercent;
            ImageDataHex = imageDataHex;
        } // RtfVisualImage

        public RtfVisualImageFormat Format { get; } // Format

        public RtfTextAlignment Alignment { get; set; } // Alignment

        public int Width { get; } // Width

        public int Height { get; } // Height

        public int DesiredWidth { get; } // DesiredWidth

        public int DesiredHeight { get; } // DesiredHeight

        public int ScaleWidthPercent { get; } // ScaleWidthPercent

        public int ScaleHeightPercent { get; } // ScaleHeightPercent

        public string ImageDataHex { get; } // ImageDataHex

        public byte[] ImageDataBinary
        {
            get { return imageDataBinary ?? (imageDataBinary = ToBinary(ImageDataHex)); }
        } // ImageDataBinary

        public Image ImageForDrawing
        {
            get
            {
                switch (Format)
                {
                    case RtfVisualImageFormat.Bmp:
                    case RtfVisualImageFormat.Jpg:
                    case RtfVisualImageFormat.Png:
                    case RtfVisualImageFormat.Emf:
                    case RtfVisualImageFormat.Wmf:
                        var data = ImageDataBinary;
                        return Image.FromStream(new MemoryStream(data, 0, data.Length));
                }
                return null;
            }
        } // ImageForDrawing

        protected override void DoVisit(IRtfVisualVisitor visitor)
        {
            visitor.VisitImage(this);
        } // DoVisit

        public static byte[] ToBinary(string imageDataHex)
        {
            if (imageDataHex == null)
                throw new ArgumentNullException("imageDataHex");

            var hexDigits = imageDataHex.Length;
            var dataSize = hexDigits / 2;
            var imageDataBinary = new byte[dataSize];

            var hex = new StringBuilder(2);

            var dataPos = 0;
            for (var i = 0; i < hexDigits; i++)
            {
                var c = imageDataHex[i];
                if (char.IsWhiteSpace(c))
                    continue;
                hex.Append(imageDataHex[i]);
                if (hex.Length == 2)
                {
                    imageDataBinary[dataPos] = byte.Parse(hex.ToString(), NumberStyles.HexNumber);
                    dataPos++;
                    hex.Remove(0, 2);
                }
            }

            return imageDataBinary;
        } // ToBinary

        protected override bool IsEqual(object obj)
        {
            var compare = obj as RtfVisualImage; // guaranteed to be non-null
            return
                compare != null &&
                base.IsEqual(compare) &&
                Format == compare.Format &&
                Alignment == compare.Alignment &&
                Width == compare.Width &&
                Height == compare.Height &&
                DesiredWidth == compare.DesiredWidth &&
                DesiredHeight == compare.DesiredHeight &&
                ScaleWidthPercent == compare.ScaleWidthPercent &&
                ScaleHeightPercent == compare.ScaleHeightPercent &&
                ImageDataHex.Equals(compare.ImageDataHex);
            //imageDataBinary.Equals( compare.imageDataBinary ); // cached info only
        } // IsEqual

        protected override int ComputeHashCode()
        {
            var hash = base.ComputeHashCode();
            hash = HashTool.AddHashCode(hash, Format);
            hash = HashTool.AddHashCode(hash, Alignment);
            hash = HashTool.AddHashCode(hash, Width);
            hash = HashTool.AddHashCode(hash, Height);
            hash = HashTool.AddHashCode(hash, DesiredWidth);
            hash = HashTool.AddHashCode(hash, DesiredHeight);
            hash = HashTool.AddHashCode(hash, ScaleWidthPercent);
            hash = HashTool.AddHashCode(hash, ScaleHeightPercent);
            hash = HashTool.AddHashCode(hash, ImageDataHex);
            //hash = HashTool.AddHashCode( hash, imageDataBinary ); // cached info only
            return hash;
        } // ComputeHashCode

        public override string ToString()
        {
            return "[" + Format + ": " + Alignment + ", " +
                   Width + " x " + Height + " " +
                   "(" + DesiredWidth + " x " + DesiredHeight + ") " +
                   "{" + ScaleWidthPercent + "% x " + ScaleHeightPercent + "%} " +
                   ":" + ImageDataHex.Length / 2 + " bytes]";
        } // ToString
    } // class RtfVisualImage
}