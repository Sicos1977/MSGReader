// -- FILE ------------------------------------------------------------------
// name       : Strings.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2009.02.19
// language   : c#
// environment: .NET 3.5
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

using System.Resources;

namespace Itenso.Sys
{
    // ------------------------------------------------------------------------
    /// <summary>Provides strongly typed resource access for this namespace.</summary>
    internal sealed class Strings : StringsBase
    {
        // ----------------------------------------------------------------------
        // members
        private static readonly ResourceManager inst = NewInst(typeof(Strings));

        // ----------------------------------------------------------------------
        public static string ArgumentMayNotBeEmpty
        {
            get { return inst.GetString("ArgumentMayNotBeEmpty"); }
        } // ArgumentMayNotBeEmpty

        // ----------------------------------------------------------------------
        public static string LoggerNameMayNotBeEmpty
        {
            get { return inst.GetString("LoggerNameMayNotBeEmpty"); }
        } // LoggerNameMayNotBeEmpty

        // ----------------------------------------------------------------------
        public static string LoggerFactoryConfigError
        {
            get { return inst.GetString("LoggerFactoryConfigError"); }
        } // LoggerFactoryConfigError

        // ----------------------------------------------------------------------
        public static string ProgramPressAnyKeyToQuit
        {
            get { return inst.GetString("ProgramPressAnyKeyToQuit"); }
        } // ProgramPressAnyKeyToQuit

        // ----------------------------------------------------------------------
        public static string StringToolSeparatorIncludesQuoteOrEscapeChar
        {
            get { return inst.GetString("StringToolSeparatorIncludesQuoteOrEscapeChar"); }
        } // StringToolSeparatorIncludesQuoteOrEscapeChar

        // ----------------------------------------------------------------------
        public static string StringToolMissingEscapedHexCode
        {
            get { return inst.GetString("StringToolMissingEscapedHexCode"); }
        } // StringToolMissingEscapedHexCode

        // ----------------------------------------------------------------------
        public static string StringToolMissingEscapedChar
        {
            get { return inst.GetString("StringToolMissingEscapedChar"); }
        } // StringToolMissingEscapedChar

        // ----------------------------------------------------------------------
        public static string StringToolUnbalancedQuotes
        {
            get { return inst.GetString("StringToolUnbalancedQuotes"); }
        } // StringToolUnbalancedQuotes

        // ----------------------------------------------------------------------
        public static string StringToolContainsInvalidHexChar
        {
            get { return inst.GetString("StringToolContainsInvalidHexChar"); }
        } // StringToolContainsInvalidHexChar

        // ----------------------------------------------------------------------
        public static string LoggerLoggingLevelXmlError
        {
            get { return inst.GetString("LoggerLoggingLevelXmlError"); }
        } // LoggerLoggingLevelXmlError

        // ----------------------------------------------------------------------
        public static string LoggerLoggingLevelRepository
        {
            get { return inst.GetString("LoggerLoggingLevelRepository"); }
        } // LoggerLoggingLevelRepository

        // ----------------------------------------------------------------------
        public static string CollectionToolInvalidEnum(string value, string enumType, string possibleValues)
        {
            return Format(inst.GetString("CollectionToolInvalidEnum"), value, enumType, possibleValues);
        } // CollectionToolInvalidEnum

        // ----------------------------------------------------------------------
        public static string LoggerLogFileNotSupportedByType(string typeName)
        {
            return Format(inst.GetString("LoggerLogFileNotSupportedByType"), typeName);
        } // LoggerLogFileNotSupportedByType
    } // class Strings
} // namespace Itenso.Sys
// -- EOF -------------------------------------------------------------------