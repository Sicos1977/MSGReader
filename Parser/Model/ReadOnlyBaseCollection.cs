// -- FILE ------------------------------------------------------------------
// name       : ReadOnlyBaseCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.20
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System.Collections;
using Itenso.Sys;
using Itenso.Sys.Collection;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public abstract class ReadOnlyBaseCollection : ReadOnlyCollectionBase
	{

		// ----------------------------------------------------------------------
		public sealed override bool Equals( object obj )
		{
			if ( obj == this )
			{
				return true;
			}
			
			if ( obj == null || GetType() != obj.GetType() )
			{
				return false;
			}

			return IsEqual( obj );
		} // Equals

		// ----------------------------------------------------------------------
		public override string ToString()
		{
			return CollectionTool.ToString( this );
		} // ToString

		// ----------------------------------------------------------------------
		protected virtual bool IsEqual( object obj )
		{
			return CollectionTool.AreEqual( this, obj );
		} // IsEqual

		// ----------------------------------------------------------------------
		public sealed override int GetHashCode()
		{
			return HashTool.AddHashCode( GetType().GetHashCode(), ComputeHashCode() );
		} // GetHashCode

		// ----------------------------------------------------------------------
		protected virtual int ComputeHashCode()
		{
			return HashTool.ComputeHashCode( this );
		} // ComputeHashCode

	} // class ReadOnlyBaseCollection

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
