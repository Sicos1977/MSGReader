// -- FILE ------------------------------------------------------------------
// name       : HelpModeArgument.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
namespace Itenso.Sys.Application
{

	// ------------------------------------------------------------------------
	public class HelpModeArgument : ToggleArgument
	{

		// ----------------------------------------------------------------------
		public const string HelpModeArgumentName = "?";

		// ----------------------------------------------------------------------
		public HelpModeArgument() :
			base( HelpModeArgumentName, false )
		{
		} // HelpModeArgument

	} // class HelpModeArgument

} // namespace Itenso.Sys.Application
// -- EOF -------------------------------------------------------------------
