// name       : IRtfParser.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;
using System.IO;

namespace Itenso.Rtf
{
    public interface IRtfParser
    {
        /// <summary>
        ///     Determines whether to ignore all content after the root group ends.
        ///     Set this to true when parsing content from streams which contain other
        ///     data after the RTF or if the writer of the RTF is known to terminate the
        ///     actual RTF content with a null byte (as some popular sources such as
        ///     WordPad are known to behave).
        /// </summary>
        bool IgnoreContentAfterRootGroup { get; set; }

        /// <summary>
        ///     Adds a listener that will get notified along the parsing process.
        /// </summary>
        /// <param name="listener">the listener to add</param>
        /// <exception cref="ArgumentNullException">in case of a null argument</exception>
        void AddParserListener(IRtfParserListener listener);

        /// <summary>
        ///     Removes a listener from this instance.
        /// </summary>
        /// <param name="listener">the listener to remove</param>
        /// <exception cref="ArgumentNullException">in case of a null argument</exception>
        void RemoveParserListener(IRtfParserListener listener);

        /// <summary>
        ///     Parses the given RTF text that is read from the given source.
        /// </summary>
        /// <param name="rtfTextSource">the source with RTF text to parse</param>
        /// <exception cref="RtfException">in case of invalid RTF syntax</exception>
        /// <exception cref="IOException">in case of an IO error</exception>
        /// <exception cref="ArgumentNullException">in case of a null argument</exception>
        void Parse(IRtfSource rtfTextSource);
    }
}