// name       : ToggleArgument.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2008.06.03
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland

using System;

namespace Itenso.Sys.Application
{
    public class ToggleArgument : Argument
    {
        public new bool Value
        {
            get { return (bool) base.Value; }
        } // Value

        public ToggleArgument(string name, bool defaultValue) :
            this(ArgumentType.None, name, defaultValue)
        {
        } // ToggleArgument

        public ToggleArgument(ArgumentType argumentType, string name, bool defaultValue) :
            base(argumentType | ArgumentType.HasName | ArgumentType.ContainsValue, name, defaultValue)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));
        } // ToggleArgument

        public override string ToString()
        {
            return Name + "=" + (Value ? "on" : "off");
        } // ToString

        protected override bool OnLoad(string commandLineArg)
        {
            if (!commandLineArg.StartsWith(Name, StringComparison.InvariantCultureIgnoreCase))
                return false;

            // format: /name
            if (commandLineArg.Equals(Name, StringComparison.InvariantCultureIgnoreCase))
            {
                base.Value = true;
                return true;
            }

            // format: /name+
            var onName = Name + "+";
            if (commandLineArg.Equals(onName, StringComparison.InvariantCultureIgnoreCase))
            {
                base.Value = true;
                return true;
            }

            // format: /name-
            var offName = Name + "-";
            if (commandLineArg.Equals(offName, StringComparison.InvariantCultureIgnoreCase))
            {
                base.Value = false;
                return true;
            }

            return false;
        } // OnLoad
    }
}