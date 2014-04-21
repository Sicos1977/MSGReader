using System;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal partial class Storage
    {
        /// <summary>
        /// Class used to contain all the contact information
        /// </summary>
        public sealed class Contact : Storage
        {
            #region Properties
            /// <summary>
            /// Returns the start datetime of the task, null when not available
            /// </summary>
            public DateTime? StartDate
            {
                get { return GetMapiPropertyDateTime(MapiTags.TaskStartDate); }
            }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Contact" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Contact(Storage message) : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _namedProperties = message._namedProperties;
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
            }
            #endregion
        }
    }
}
