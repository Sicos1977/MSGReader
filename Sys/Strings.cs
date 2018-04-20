// name       : Strings.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2009.02.19
// language   : c#
// environment: .NET 3.5
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System.Resources;

namespace Itenso.Sys
{
    /// <summary>Provides strongly typed resource access for this namespace.</summary>
    internal sealed class Strings : StringsBase
    {
        // Members
        private static readonly ResourceManager Inst = NewInst(typeof(Strings));

        public static string ArgumentMayNotBeEmpty
        {
            get { return Inst.GetString("ArgumentMayNotBeEmpty"); }
        } // ArgumentMayNotBeEmpty

        public static string LoggerNameMayNotBeEmpty
        {
            get { return Inst.GetString("LoggerNameMayNotBeEmpty"); }
        } // LoggerNameMayNotBeEmpty

        public static string LoggerFactoryConfigError
        {
            get { return Inst.GetString("LoggerFactoryConfigError"); }
        } // LoggerFactoryConfigError

        public static string ProgramPressAnyKeyToQuit
        {
            get { return Inst.GetString("ProgramPressAnyKeyToQuit"); }
        } // ProgramPressAnyKeyToQuit

        public static string StringToolSeparatorIncludesQuoteOrEscapeChar
        {
            get { return Inst.GetString("StringToolSeparatorIncludesQuoteOrEscapeChar"); }
        } // StringToolSeparatorIncludesQuoteOrEscapeChar

        public static string StringToolMissingEscapedHexCode
        {
            get { return Inst.GetString("StringToolMissingEscapedHexCode"); }
        } // StringToolMissingEscapedHexCode

        public static string StringToolMissingEscapedChar
        {
            get { return Inst.GetString("StringToolMissingEscapedChar"); }
        } // StringToolMissingEscapedChar

        public static string StringToolUnbalancedQuotes
        {
            get { return Inst.GetString("StringToolUnbalancedQuotes"); }
        } // StringToolUnbalancedQuotes

        public static string StringToolContainsInvalidHexChar
        {
            get { return Inst.GetString("StringToolContainsInvalidHexChar"); }
        } // StringToolContainsInvalidHexChar

        public static string LoggerLoggingLevelXmlError
        {
            get { return Inst.GetString("LoggerLoggingLevelXmlError"); }
        } // LoggerLoggingLevelXmlError

        public static string LoggerLoggingLevelRepository
        {
            get { return Inst.GetString("LoggerLoggingLevelRepository"); }
        } // LoggerLoggingLevelRepository

        public static string CollectionToolInvalidEnum(string value, string enumType, string possibleValues)
        {
            return Format(Inst.GetString("CollectionToolInvalidEnum"), value, enumType, possibleValues);
        } // CollectionToolInvalidEnum

        public static string LoggerLogFileNotSupportedByType(string typeName)
        {
            return Format(Inst.GetString("LoggerLogFileNotSupportedByType"), typeName);
        } // LoggerLogFileNotSupportedByType
    }
}