using System;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class used to contain all the contact information
        /// </summary>
        public sealed class Contact : Storage
        {
            #region Properties
            /// <summary>
            /// Returns the full name (e.g. De heer Kees van Spelde), null when not available
            /// </summary>
            public string DisplayName { get; private set; }

            /// <summary>
            /// Returns the prefix of the name (e.g. De heer), null when not available
            /// </summary>
            public string Prefix { get; private set; }

            /// <summary>
            /// Returns the initials, null when not available
            /// </summary>
            public string Initials { get; private set; }

            /// <summary>
            /// Returns the sur name (e.g. Spelde), null when not available
            /// </summary>
            public string SurName { get; private set; }

            /// <summary>
            /// Returns the given name (e.g. Kees), null when not available
            /// </summary>
            public string GivenName { get; private set; }

            /// <summary>
            /// Returns the generation (e.g. Jr.)
            /// </summary>
            public string Generation { get; private set; }

            /// <summary>
            /// Returns the function, null when not available
            /// </summary>
            public string Function { get; private set; }

            /// <summary>
            /// Returns the department, null when not available
            /// </summary>
            public string Department { get; private set; }

            /// <summary>
            /// Returns the name of the company, null when not available
            /// </summary>
            public string Company { get; private set; }
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

                DisplayName = GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME);
                Prefix = GetMapiPropertyString(MapiTags.PR_DISPLAY_NAME_PREFIX);
                Initials = GetMapiPropertyString(MapiTags.PR_INITIALS);
                SurName = GetMapiPropertyString(MapiTags.PR_SURNAME);
                GivenName = GetMapiPropertyString(MapiTags.PR_GIVEN_NAME);
                Generation = GetMapiPropertyString(MapiTags.PR_GENERATION);
                Function = GetMapiPropertyString(MapiTags.PR_TITLE);
                Department = GetMapiPropertyString(MapiTags.PR_DEPARTMENT_NAME);
                Company = GetMapiPropertyString(MapiTags.PR_COMPANY_NAME);
            }
            #endregion
        }
    }
}
