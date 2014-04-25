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

            #region Work address information
            /// <summary>
            /// Returns the street of the work address
            /// </summary>
            public string WorkAddressStreet { get; private set; }

            /// <summary>
            /// Returns the city of the work address
            /// </summary>
            public string WorkAddressCity { get; private set; }

            /// <summary>
            /// Returns the state of the work address
            /// </summary>
            public string WorkAddressState { get; private set; }

            /// <summary>
            /// Returns the postal code of the work address
            /// </summary>
            public string WorkAddressPostalCode { get; private set; }

            /// <summary>
            /// Returns the country of the work address
            /// </summary>
            public string WorkAddressCountry { get; private set; }
            #endregion

            #region Home address information
            /// <summary>
            /// Returns the street of the home address
            /// </summary>
            public string HomeAddressStreet { get; private set; }

            /// <summary>
            /// Returns the city of the home address
            /// </summary>
            public string HomeAddressCity { get; private set; }

            /// <summary>
            /// Returns the state of the home address
            /// </summary>
            public string HomeAddressState { get; private set; }

            /// <summary>
            /// Returns the postal code of the home address
            /// </summary>
            public string HomeAddressPostalCode { get; private set; }

            /// <summary>
            /// Returns the country of the home address
            /// </summary>
            public string HomeAddressCountry { get; private set; }
            #endregion
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

                // Work address information
                WorkAddressStreet  = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_STREET);
                WorkAddressCity = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_CITY);
                WorkAddressState = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE);
                WorkAddressPostalCode = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_POSTAL_CODE);
                WorkAddressCountry = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_COUNTRY);

                // Home address information
                HomeAddressStreet = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_STREET);
                HomeAddressCity = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_CITY);
                HomeAddressState = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_STATE_OR_PROVINCE);
                HomeAddressPostalCode = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_POSTAL_CODE);
                HomeAddressCountry = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_COUNTRY);
            }
            #endregion
        }
    }
}
