// -- FILE ------------------------------------------------------------------
// name       : IRtfParserListener.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Rtf
{
    // ------------------------------------------------------------------------
    public interface IRtfParserListener
    {
        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called before any other of the methods upon starting parsing of new input.
        /// </summary>
        void ParseBegin();

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called when a new group began.
        /// </summary>
        void GroupBegin();

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called when a new tag was found.
        /// </summary>
        /// <param name="tag">the newly found tag</param>
        void TagFound(IRtfTag tag);

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called when a new text was found.
        /// </summary>
        /// <param name="text">the newly found text</param>
        void TextFound(IRtfText text);

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called after a group ended.
        /// </summary>
        void GroupEnd();

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called if parsing finished sucessfully.
        /// </summary>
        void ParseSuccess();

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called if parsing failed.
        /// </summary>
        /// <param name="reason">the reason for the failure</param>
        void ParseFail(RtfException reason);

        // ----------------------------------------------------------------------
        /// <summary>
        ///     Called after parsing finished. Always called, also in case of a failure.
        /// </summary>
        void ParseEnd();
    } // interface IRtfParserListener
} // namespace Itenso.Rtf
// -- EOF -------------------------------------------------------------------