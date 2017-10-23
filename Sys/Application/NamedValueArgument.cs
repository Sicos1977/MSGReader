// -- FILE ------------------------------------------------------------------
// name       : NamedValueArgument.cs
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
	public class NamedValueArgument : Argument
	{

		// ----------------------------------------------------------------------
		public NamedValueArgument( string name ) :
			this( ArgumentType.None, name, null )
		{
		} // NamedValueArgument

		// ----------------------------------------------------------------------
		public NamedValueArgument( string name, object defaultValue ) :
			this( ArgumentType.None, name, defaultValue )
		{
		} // NamedValueArgument

		// ----------------------------------------------------------------------
		public NamedValueArgument( ArgumentType argumentType, string name, object defaultValue ) :
			base( argumentType | ArgumentType.ContainsValue | ArgumentType.HasName, name, defaultValue )
		{
		} // NamedValueArgument

		// ----------------------------------------------------------------------
		public new string Value
		{
			get { return base.Value as string; }
		} // Value

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return Name + "=" + Value;
		} // ToString

		// ----------------------------------------------------------------------
		protected override bool OnLoad( string commandLineArg )
		{
			// format: /name:value
			string valueName = Name + ":";
			if ( !commandLineArg.StartsWith( valueName, StringComparison.InvariantCultureIgnoreCase ) )
			{
				return false;
			}

			base.Value = commandLineArg.Substring( valueName.Length );
			return true;
		} // OnLoad

	} // class NamedValueArgument

} // namespace Itenso.Sys.Application
// -- EOF -------------------------------------------------------------------
