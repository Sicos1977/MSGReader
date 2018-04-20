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
        // members
        private object value;

        // ----------------------------------------------------------------------
        public ArgumentType ArgumentType { get; } // ArgumentType

        // ----------------------------------------------------------------------
        public bool ContainsValue
        {
            get { return (ArgumentType & ArgumentType.ContainsValue) == ArgumentType.ContainsValue; }
        } // ContainsValue

        // ----------------------------------------------------------------------
        protected Argument(ArgumentType argumentType, string name, object defaultValue)
        {
            Name = name;
            ArgumentType = argumentType;
            DefaultValue = defaultValue;
        } // Argument

        // ----------------------------------------------------------------------
        public string Name { get; } // Name

        // ----------------------------------------------------------------------
        public object Value
        {
            get
            {
                if (value == null)
                    return DefaultValue;
                return value;
            }
            set { this.value = value; }
        } // Value

        // ----------------------------------------------------------------------
        public object DefaultValue { get; } // DefaultValue

        // ----------------------------------------------------------------------
        public bool IsMandatory
        {
            get { return (ArgumentType & ArgumentType.Mandatory) == ArgumentType.Mandatory; }
        } // IsMandatory

        // ----------------------------------------------------------------------
        public bool HasName
        {
            get { return (ArgumentType & ArgumentType.HasName) == ArgumentType.HasName; }
        } // HasName

        // ----------------------------------------------------------------------
        public virtual bool IsValid
        {
            get
            {
                if (IsMandatory && !IsLoaded)
                    return false;

                if (IsMandatory && ContainsValue && Value == null && DefaultValue == null)
                    return false;

                return true;
            }
        } // IsValid

        // ----------------------------------------------------------------------
        public bool IsLoaded { get; private set; } // IsLoaded

        // ----------------------------------------------------------------------
        public void Load(string commandLineArg)
        {
            var isNamedArg = commandLineArg.StartsWith("/") || commandLineArg.StartsWith("-");

            // argument with name
            if (HasName)
            {
                if (!isNamedArg)
                    return; // missing argument name

                commandLineArg = commandLineArg.Substring(1);
                if (string.IsNullOrEmpty(commandLineArg))
                    return;
            }
            else if (isNamedArg)
            {
                return; // name provided on argument without name
            }

            IsLoaded = OnLoad(commandLineArg);
        } // Load

        // ----------------------------------------------------------------------
        protected abstract bool OnLoad(string commandLineArg);
    } // class ArgumentInfo
} // namespace Itenso.Sys.Application
// -- EOF -------------------------------------------------------------------