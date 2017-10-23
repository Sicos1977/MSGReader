// -- FILE ------------------------------------------------------------------
// name       : ArgumentInfo.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
namespace Itenso.Sys.Application
{

	// ------------------------------------------------------------------------
	public abstract class Argument : IArgument
	{

		// ----------------------------------------------------------------------
		protected Argument( ArgumentType argumentType, string name, object defaultValue )
		{
			this.name = name;
			this.argumentType = argumentType;
			this.defaultValue = defaultValue;
		} // Argument

		// ----------------------------------------------------------------------
		public string Name
		{
			get { return name; }
		} // Name

		// ----------------------------------------------------------------------
		public object Value
		{
			get 
			{
				if ( value == null )
				{
					return defaultValue;
				}
				return value; 
			}
			set { this.value = value; }
		} // Value

		// ----------------------------------------------------------------------
		public object DefaultValue
		{
			get { return defaultValue; }
		} // DefaultValue

		// ----------------------------------------------------------------------
		public ArgumentType ArgumentType
		{
			get { return argumentType; }
		} // ArgumentType

		// ----------------------------------------------------------------------
		public bool IsMandatory
		{
			get { return ( argumentType & ArgumentType.Mandatory ) == ArgumentType.Mandatory; }
		} // IsMandatory

		// ----------------------------------------------------------------------
		public bool HasName
		{
			get { return ( argumentType & ArgumentType.HasName ) == ArgumentType.HasName; }
		} // HasName

		// ----------------------------------------------------------------------
		public bool ContainsValue
		{
			get { return ( argumentType & ArgumentType.ContainsValue ) == ArgumentType.ContainsValue; }
		} // ContainsValue

		// ----------------------------------------------------------------------
		public virtual bool IsValid
		{
			get 
			{
				if ( IsMandatory && !isLoaded )
				{
					return false;
				}

				if ( IsMandatory && ContainsValue && ( Value == null && DefaultValue == null ) )
				{
					return false;
				}

				return true; 
			}
		} // IsValid

		// ----------------------------------------------------------------------
		public bool IsLoaded
		{
			get { return isLoaded; }
		} // IsLoaded

		// ----------------------------------------------------------------------
		public void Load( string commandLineArg )
		{
			bool isNamedArg = commandLineArg.StartsWith( "/" ) || commandLineArg.StartsWith( "-" );

			// argument with name
			if ( HasName )
			{
				if ( !isNamedArg )
				{
					return; // missing argument name
				}

				commandLineArg = commandLineArg.Substring( 1 );
				if ( string.IsNullOrEmpty( commandLineArg ) )
				{
					return;
				}
			}
			else if ( isNamedArg )
			{
				return; // name provided on argument without name
			}

			isLoaded = OnLoad( commandLineArg );
		} // Load

		// ----------------------------------------------------------------------
		protected abstract bool OnLoad( string commandLineArg );

		// ----------------------------------------------------------------------
		// members
		private readonly string name;
		private readonly ArgumentType argumentType;
		private readonly object defaultValue;
		private bool isLoaded;
		private object value;

	} // class ArgumentInfo

} // namespace Itenso.Sys.Application
// -- EOF -------------------------------------------------------------------
