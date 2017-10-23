// -- FILE ------------------------------------------------------------------
// name       : RtfTag.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Globalization;
using Itenso.Sys;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfTag : RtfElement, IRtfTag
	{

		// ----------------------------------------------------------------------
		public RtfTag( string name ) :
			base( RtfElementKind.Tag )
		{
			if ( name == null )
			{
				throw new ArgumentNullException( "name" );
			}
			fullName = name;
			this.name = name;
			valueAsText = null;
			valueAsNumber = -1;
		} // RtfTag

		// ----------------------------------------------------------------------
		public RtfTag( string name, string value ) :
			base( RtfElementKind.Tag )
		{
			if ( name == null )
			{
				throw new ArgumentNullException( "name" );
			}
			if ( value == null )
			{
				throw new ArgumentNullException( "value" );
			}
			fullName = name + value;
			this.name = name;
			valueAsText = value;
			int numericalValue;
			if ( Int32.TryParse( value, NumberStyles.Integer, CultureInfo.InvariantCulture, out numericalValue ) )
			{
				valueAsNumber = numericalValue;
			}
			else
			{
				valueAsNumber = -1;
			}
		} // RtfTag

		// ----------------------------------------------------------------------
		public string FullName
		{
			get { return fullName; }
		} // FullName

		// ----------------------------------------------------------------------
		public string Name
		{
			get { return name; }
		} // Name

		// ----------------------------------------------------------------------
		public bool HasValue
		{
			get { return valueAsText != null; }
		} // HasValue

		// ----------------------------------------------------------------------
		public string ValueAsText
		{
			get { return valueAsText; }
		} // ValueAsText

		// ----------------------------------------------------------------------
		public int ValueAsNumber
		{
			get { return valueAsNumber; }
		} // ValueAsNumber

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return "\\" + fullName;
		} // ToString

		// ----------------------------------------------------------------------
		protected override void DoVisit( IRtfElementVisitor visitor )
		{
			visitor.VisitTag( this );
		} // DoVisit

		// ----------------------------------------------------------------------
		protected override bool IsEqual( object obj )
		{
			RtfTag compare = obj as RtfTag; // guaranteed to be non-null
			return compare != null && base.IsEqual( obj ) &&
				fullName.Equals( compare.fullName );
		} // IsEqual

		// ----------------------------------------------------------------------
		protected override int ComputeHashCode()
		{
			int hash = base.ComputeHashCode();
			hash = HashTool.AddHashCode( hash, fullName );
			return hash;
		} // ComputeHashCode

		// ----------------------------------------------------------------------
		// members
		private readonly string fullName;
		private readonly string name;
		private readonly string valueAsText;
		private readonly int valueAsNumber;

	} // class RtfTag

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
