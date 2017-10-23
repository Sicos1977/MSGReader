// -- FILE ------------------------------------------------------------------
// name       : ValueArgument.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------

namespace Itenso.Sys.Application
{

	// ------------------------------------------------------------------------
	public class ValueArgument : Argument
	{

		// ----------------------------------------------------------------------
		public ValueArgument() :
			this( ArgumentType.None )
		{
		} // ValueArgument

		// ----------------------------------------------------------------------
		public ValueArgument( ArgumentType argumentType ) :
			base( argumentType | ArgumentType.ContainsValue, null, null )
		{
		} // ValueArgument

		// ----------------------------------------------------------------------
		public new string Value
		{
			get { return base.Value as string; }
		} // Value

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return Value;
		} // ToString

		// ----------------------------------------------------------------------
		protected override bool OnLoad( string commandLineArg )
		{
			base.Value = commandLineArg;
			return true;
		} // OnLoad

	} // class ValueArgument

} // namespace Itenso.Sys.Application
// -- EOF -------------------------------------------------------------------
