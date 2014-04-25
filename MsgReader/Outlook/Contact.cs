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
            /// Returns the generation (e.g. Jr.), null when not available
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

            #region Business address information
            /// <summary>
            /// Returns the street of the business address, null when not available
            /// </summary>
            public string BusinessAddressStreet { get; private set; }

            /// <summary>
            /// Returns the city of the business address, null when not available
            /// </summary>
            public string BusinessAddressCity { get; private set; }

            /// <summary>
            /// Returns the state of the business address, null when not available
            /// </summary>
            public string BusinessAddressState { get; private set; }

            /// <summary>
            /// Returns the postal code of the business address, null when not available
            /// </summary>
            public string BusinessAddressPostalCode { get; private set; }

            /// <summary>
            /// Returns the country of the business address, null when not available
            /// </summary>
            public string BusinessAddressCountry { get; private set; }

            /// <summary>
            /// Returns the business telephone number, null when not available
            /// </summary>
            public string BusinessTelephoneNumber { get; private set; }

            /// <summary>
            /// Returns the business second telephone number, null when not available
            /// </summary>
            public string BusinessTelephoneNumber2 { get; private set; }
            
            /// <summary>
            /// Returns the business fax number, null when not available
            /// </summary>
            public string BusinessFaxNumber { get; private set; }

            /// <summary>
            /// Returns the business home page, null when not available
            /// </summary>
            public string BusinessHomePage { get; private set; }
            #endregion

            #region Home address information
            /// <summary>
            /// Returns the street of the home address, null when not available
            /// </summary>
            public string HomeAddressStreet { get; private set; }

            /// <summary>
            /// Returns the city of the home address, null when not available
            /// </summary>
            public string HomeAddressCity { get; private set; }

            /// <summary>
            /// Returns the state of the home address, null when not available
            /// </summary>
            public string HomeAddressState { get; private set; }

            /// <summary>
            /// Returns the postal code of the home address, null when not available
            /// </summary>
            public string HomeAddressPostalCode { get; private set; }

            /// <summary>
            /// Returns the country of the home address, null when not available
            /// </summary>
            public string HomeAddressCountry { get; private set; }

            /// <summary>
            /// Returns the home telephone number, null when not available
            /// </summary>
            public string HomeTelephoneNumber { get; private set; }

            /// <summary>
            /// Returns the home second telephone number, null when not available
            /// </summary>
            public string HomeTelephoneNumber2 { get; private set; }

            /// <summary>
            /// Returns the home fax number, null when not available
            /// </summary>
            public string HomeFaxNumber { get; private set; }
            #endregion

            #region Other address information
            /// <summary>
            /// Returns the street of the other address, null when not available
            /// </summary>
            public string OtherAddressStreet { get; private set; }

            /// <summary>
            /// Returns the city of the other address, null when not available
            /// </summary>
            public string OtherAddressCity { get; private set; }

            /// <summary>
            /// Returns the state of the other address, null when not available
            /// </summary>
            public string OtherAddressState { get; private set; }

            /// <summary>
            /// Returns the postal code of the other address, null when not available
            /// </summary>
            public string OtherAddressPostalCode { get; private set; }

            /// <summary>
            /// Returns the country of the other address, null when not available
            /// </summary>
            public string OtherAddressCountry { get; private set; }
            #endregion

            /// <summary>
            /// Return the instant messaging address, null when not available
            /// </summary>
            public string InstantMessagingAddress { get; private set; }
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
                BusinessAddressStreet  = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_STREET);
                BusinessAddressCity = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_CITY);
                BusinessAddressState = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE);
                BusinessAddressPostalCode = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_POSTAL_CODE);
                BusinessAddressCountry = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_COUNTRY);
                BusinessTelephoneNumber = GetMapiPropertyString(MapiTags.PR_BUSINESS_TELEPHONE_NUMBER);
                BusinessTelephoneNumber2 = GetMapiPropertyString(MapiTags.PR_BUSINESS2_TELEPHONE_NUMBER);
                BusinessFaxNumber = GetMapiPropertyString(MapiTags.PR_BUSINESS_FAX_NUMBER);
                BusinessHomePage = GetMapiPropertyString(MapiTags.PR_BUSINESS_HOME_PAGE);

                // Home address information
                HomeAddressStreet = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_STREET);
                HomeAddressCity = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_CITY);
                HomeAddressState = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_STATE_OR_PROVINCE);
                HomeAddressPostalCode = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_POSTAL_CODE);
                HomeAddressCountry = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_COUNTRY);
                HomeTelephoneNumber = GetMapiPropertyString(MapiTags.PR_HOME_TELEPHONE_NUMBER);
                HomeTelephoneNumber2 = GetMapiPropertyString(MapiTags.PR_HOME2_TELEPHONE_NUMBER);
                HomeFaxNumber = GetMapiPropertyString(MapiTags.PR_HOME_FAX_NUMBER);
                
                // Other address information
                OtherAddressStreet = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_STREET);
                OtherAddressCity = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_CITY);
                OtherAddressState = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_STATE_OR_PROVINCE);
                OtherAddressPostalCode = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_POSTAL_CODE);
                OtherAddressCountry = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_COUNTRY);

                InstantMessagingAddress = GetMapiPropertyString(MapiTags.InstantMessagingAddress);
            }
            #endregion
        }
    }
}
