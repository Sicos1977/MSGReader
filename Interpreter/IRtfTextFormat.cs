// -- FILE ------------------------------------------------------------------
// name       : IRtfTextFormat.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{
    // ------------------------------------------------------------------------
    public interface IRtfTextFormat
    {
        // ----------------------------------------------------------------------
        IRtfFont Font { get; }

        // ----------------------------------------------------------------------
        int FontSize { get; }

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Combines the setting for sub/super script: negative values are considered
        ///     equivalent to subscript, positive values correspond to superscript.<br />
        ///     Same unit as font size.
        /// </summary>
        int SuperScript { get; }

        // ----------------------------------------------------------------------
        bool IsBold { get; }

        // ----------------------------------------------------------------------
        bool IsItalic { get; }

        // ----------------------------------------------------------------------
        bool IsUnderline { get; }

        // ----------------------------------------------------------------------
        bool IsStrikeThrough { get; }

        // ----------------------------------------------------------------------
        bool IsHidden { get; }

        // ----------------------------------------------------------------------
        string FontDescriptionDebug { get; }

        // ----------------------------------------------------------------------
        IRtfColor BackgroundColor { get; }

        // ----------------------------------------------------------------------
        IRtfColor ForegroundColor { get; }

        // ----------------------------------------------------------------------
        RtfTextAlignment Alignment { get; }

        // ----------------------------------------------------------------------
        IRtfTextFormat Duplicate();
    } // interface IRtfTextFormat
} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------