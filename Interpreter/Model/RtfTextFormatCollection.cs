// -- FILE ------------------------------------------------------------------
// name       : RtfTextFormatCollection.cs
// project    : RTF Framelet
// created    : Leon Poyyayil - 2008.05.21
// language   : c#
// environment: .NET 2.0
// copyright  : (c) 2004-2013 by Jani Giannoudis, Switzerland
// --------------------------------------------------------------------------
using System;

namespace Itenso.Rtf.Model
{

	// ------------------------------------------------------------------------
	public sealed class RtfTextFormatCollection : ReadOnlyBaseCollection, IRtfTextFormatCollection
	{

		// ----------------------------------------------------------------------
		public IRtfTextFormat this[ int index ]
		{
			get { return InnerList[ index ] as IRtfTextFormat; }
		} // this[ int ]

		// ----------------------------------------------------------------------
		public bool Contains( IRtfTextFormat format )
		{
			return IndexOf( format ) >= 0;
		} // Contains

		// ----------------------------------------------------------------------
		public int IndexOf( IRtfTextFormat format )
		{
			if ( format != null )
			{
				// PERFORMANCE: most probably we should maintain a hashmap for fast searching ...
				int count = Count;
				for ( int i = 0; i < count; i++ )
				{
					if ( format.Equals( InnerList[ i ] ) )
					{
						return i;
					}
				}
			}
			return -1;
		} // IndexOf

		// ----------------------------------------------------------------------
		public void CopyTo( IRtfTextFormat[] array, int index )
		{
			InnerList.CopyTo( array, index );
		} // CopyTo

		// ----------------------------------------------------------------------
		public void Add( IRtfTextFormat item )
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

	} // class RtfTextFormatCollection

} // namespace Itenso.Rtf.Model
// -- EOF -------------------------------------------------------------------
