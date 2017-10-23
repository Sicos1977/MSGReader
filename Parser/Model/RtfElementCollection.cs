// -- FILE ------------------------------------------------------------------
// name       : RtfElementCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.19
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfElementCollection : ReadOnlyBaseCollection, IRtfElementCollection
	{

		// ----------------------------------------------------------------------
		public IRtfElement this[ int index ]
		{
			get { return InnerList[ index ] as IRtfElement; }
		} // this[ int ]

		// ----------------------------------------------------------------------
		public void CopyTo( IRtfElement[] array, int index )
		{
			InnerList.CopyTo( array, index );
		} // CopyTo

		// ----------------------------------------------------------------------
		public void Add( IRtfElement item )
		{
			if ( item == null )
			{
				throw new ArgumentNullException( "item" );
			}
			InnerList.Add( item );
		} // Add

		// ----------------------------------------------------------------------
		public void Clear()
		{
			InnerList.Clear();
		} // Clear

	} // class RtfElementCollection

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
