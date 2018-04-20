// name       : ApplicationArguments.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Sys.Application
{
    public class ApplicationArguments
    {
        // Members

        public ArgumentCollection Arguments { get; } = new ArgumentCollection();

// Arguments

        public bool IsValid
        {
            get { return Arguments.IsValid; }
        } // IsValid

        public bool IsHelpMode
        {
            get
            {
                foreach (IArgument argument in Arguments)
                    if (argument is HelpModeArgument)
                        return (bool) argument.Value;
                return false;
            }
        } // IsHelpMode

        public void Load()
        {
            var commandLineArgs = Environment.GetCommandLineArgs();

            // skip zeron index which contians the program name
            for (var i = 1; i < commandLineArgs.Length; i++)
            {
                var commandLineArg = commandLineArgs[i];
                foreach (IArgument argument in Arguments)
                {
                    if (argument.IsLoaded)
                        continue;
                    argument.Load(commandLineArg);
                    if (argument.IsLoaded)
                        break;
                }
            }
        } // Load
    }
}