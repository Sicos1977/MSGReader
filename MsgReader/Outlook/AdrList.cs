using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MsgReader.Outlook
{
    /// <summary>
    /// Describes zero or more properties belonging to one or more recipients. 
    /// </summary>
    /// <remarks>
    /// https://msdn.microsoft.com/en-us/library/ms629442%28VS.85%29.aspx
    /// </remarks>
    internal class AdrList : List<AdrEntry>
    {
    }

    /// <summary>
    /// Describes zero or more properties belonging to one or more recipients.
    /// </summary>
    internal class AdrEntry
    {
        /// <summary>
        ///     Reserved.Must be zero
        /// </summary>
        public ulong Reserved1 { get; set; }

        /// <summary>
        ///     Variable of type ULONG that specifies the count of properties 
        ///     in the property value array to which the rgPropVals member points.The cValues member can be zero.
        /// </summary>
        public ulong Values { get; set; }

        /// <summary>
        ///     Pointer to a variable of type SPropValue that specifies the property value array describing the 
        ///     properties for the recipient. The rgPropVals member can be NULL.
        /// </summary>
        public SPropValue PropValue {get; set; }
    }

    /// <summary>
    ///     Contains the property tag values.
    /// </summary>
    /// <remarks>
    ///     https://msdn.microsoft.com/en-us/library/ms629450%28v=vs.85%29.aspx
    /// </remarks>
    internal class SPropValue
    {
        
    }
}
