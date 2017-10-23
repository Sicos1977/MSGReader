// -- FILE ------------------------------------------------------------------
// name       : HashTool.cs
// project    : System Framelet
// created    : Jani Giannoudis - 2005.12.01
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;
using System.Collections;

namespace Itenso.Sys
{

	// ------------------------------------------------------------------------
	/// <summary>
	/// Some hash utility methods for collections.
	/// </summary>
	public static class HashTool
	{

		// ----------------------------------------------------------------------
		public static int AddHashCode( int hash, object obj )
		{
			int combinedHash = obj != null ? obj.GetHashCode() : 0;
			if ( hash != 0 ) // perform this check to prevent FxCop warning 'op could overflow'
			{
				combinedHash += hash * 31;
			}
			return combinedHash;
		} // AddHashCode

		// ----------------------------------------------------------------------
		public static int AddHashCode( int hash, int objHash )
		{
			int combinedHash = objHash;
			if ( hash != 0 ) // perform this check to prevent FxCop warning 'op could overflow'
			{
				combinedHash += hash * 31;
			}
			return combinedHash;
		} // AddHashCode

		// ----------------------------------------------------------------------
		public static int ComputeHashCode( IEnumerable enumerable )
		{
			int hash = 1;
			if ( enumerable == null )
			{
				throw new ArgumentNullException( "enumerable" );
			}
			foreach ( object item in enumerable )
			{
				hash = hash * 31 + ( item != null ? item.GetHashCode() : 0 );
			}
			return hash;
		} // ComputeHashCode

	} // class HashTool

} // namespace Itenso.Sys
// -- EOF -------------------------------------------------------------------
