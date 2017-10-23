// -- FILE ------------------------------------------------------------------
// name       : ApplicationArguments.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;

namespace Itenso.Sys.Application
{

	// ------------------------------------------------------------------------
	public class ApplicationArguments
	{

		// ----------------------------------------------------------------------
		public ArgumentCollection Arguments
		{
			get { return arguments; }
		} // Arguments

		// ----------------------------------------------------------------------
		public bool IsValid
		{
			get { return arguments.IsValid; }
		} // IsValid

		// ----------------------------------------------------------------------
		public bool IsHelpMode
		{
			get 
			{
				foreach ( IArgument argument in arguments )
				{
					if ( argument is HelpModeArgument )
					{
						return (bool)argument.Value;
					}
				}
				return false;
			}
		} // IsHelpMode

		// ----------------------------------------------------------------------
		public void Load()
		{
			string[] commandLineArgs = Environment.GetCommandLineArgs();

			// skip zeron index which contians the program name
			for ( int i = 1; i < commandLineArgs.Length; i++ )
			{
				string commandLineArg = commandLineArgs[ i ];
				foreach ( IArgument argument in arguments )
				{
					if ( argument.IsLoaded )
					{
						continue;
					}
					argument.Load( commandLineArg );
					if ( argument.IsLoaded )
					{
						break;
					}
				}
			}
		} // Load

		// ----------------------------------------------------------------------
		// members
		private readonly ArgumentCollection arguments = new ArgumentCollection();

	} // class ApplicationArguments

} // namespace Itenso.Sys.Application
// -- EOF -------------------------------------------------------------------
