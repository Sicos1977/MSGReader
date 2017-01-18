using System;

/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader.Outlook
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
            /// Returns the named propery work (business) address (Outlook 2007 or higher), null when not available
            /// </summary>
            public string WorkAddress { get; private set; }

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
            /// Returns the named propery home address (Outlook 2007 or higher)
            /// </summary>
            public string HomeAddress { get; private set; }

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
            /// Returns the named propery other address (Outlook 2007 or higher)
            /// </summary>
            public string OtherAddress { get; private set; }

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

            /// <summary>
            /// Returns the other telephone number, null when not available
            /// </summary>
            public string OtherTelephoneNumber { get; private set; }
            #endregion

            #region Primary information
            /// <summary>
            /// Returns the primary telephone number, null when not available
            /// </summary>
            public string PrimaryTelephoneNumber { get; private set; }

            /// <summary>
            /// Returns the primary fax number, null when not available
            /// </summary>
            public string PrimaryFaxNumber { get; private set; }
            #endregion

            #region Assistant information
            /// <summary>
            /// Returns the name of the assistant, null when not available
            /// </summary>
            public string AssistantName { get; private set; }

            /// <summary>
            /// Returns the telephone number of the assistant, null when not available
            /// </summary>
            public string AssistantTelephoneNumber { get; private set; }
            #endregion
            
            /// <summary>
            /// Return the instant messaging address, null when not available
            /// </summary>
            public string InstantMessagingAddress { get; private set; }

            #region Telephone numbers
            /// <summary>
            /// Returns the company main telephone number, null when not available
            /// </summary>
            public string CompanyMainTelephoneNumber { get; private set; }
            
            /// <summary>
            /// Returns the cellular telephone number, null when not available
            /// </summary>
            public string CellularTelephoneNumber { get; private set; }
            
            /// <summary>
            /// Returns the car telephone number, null when not available
            /// </summary>
            public string CarTelephoneNumber { get; private set; }
            
            /// <summary>
            /// Returns the radio telephone number, null when not available
            /// </summary>
            public string RadioTelephoneNumber { get; private set; }

            /// <summary>
            /// Returns the beeper telephone number, null when not available
            /// </summary>
            public string BeeperTelephoneNumber { get; private set; }

            /// <summary>
            /// Returns the callback telephone number, null when not available
            /// </summary>
            public string CallbackTelephoneNumber { get; private set; }

            /// <summary>
            /// Returns the telex number, null when not available
            /// </summary>
            public string TelexNumber { get; private set; }

            /// <summary>
            /// Returns the text telephone (TTYTDD), null when not available
            /// </summary>
            public string TextTelephone { get; private set; }

            /// <summary>
            /// Returns the ISDN number, null when not available
            /// </summary>
            public string ISDNNumber { get; private set; }
            #endregion

            #region E-mail information
            /// <summary>
            /// Returns the name property e-mail address 1 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email1EmailAddress { get; private set; }

            /// <summary>
            /// Returns the name property e-mail displayname 1 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email1DisplayName { get; private set; }

            /// <summary>
            /// Returns the name property e-mail address 2 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email2EmailAddress { get; private set; }

            /// <summary>
            /// Returns the name property e-mail displayname 2 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email2DisplayName { get; private set; }
            
            /// <summary>
            /// Returns the name property e-mail address 3 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email3EmailAddress { get; private set; }

            /// <summary>
            /// Returns the name property e-mail displayname 3 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email3DisplayName { get; private set; }
            #endregion

            /// <summary>
            /// Returns the birthday, null when not available
            /// </summary>
            public DateTime? Birthday { get; private set; }

            /// <summary>
            /// Returns the wedding/anniversary, null when not available
            /// </summary>
            public DateTime? WeddingAnniversary { get; private set; }

            /// <summary>
            /// Returns the name of the spouse, null when not available
            /// </summary>
            public string SpouseName { get; private set; }

            /// <summary>
            /// Returns the profession, null when not available
            /// </summary>
            public string Profession { get; private set; }

            /// <summary>
            /// Returns the homepage (html), null when not available
            /// </summary>
            public string Html { get; private set; }
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

                #region Business address information
                WorkAddress = GetMapiPropertyString(MapiTags.WorkAddress);
                BusinessAddressStreet  = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_STREET);
                BusinessAddressCity = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_CITY);
                BusinessAddressState = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_STATE_OR_PROVINCE);
                BusinessAddressPostalCode = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_POSTAL_CODE);
                BusinessAddressCountry = GetMapiPropertyString(MapiTags.PR_BUSINESS_ADDRESS_COUNTRY);
                BusinessTelephoneNumber = GetMapiPropertyString(MapiTags.PR_BUSINESS_TELEPHONE_NUMBER);
                BusinessTelephoneNumber2 = GetMapiPropertyString(MapiTags.PR_BUSINESS2_TELEPHONE_NUMBER);
                BusinessFaxNumber = GetMapiPropertyString(MapiTags.PR_BUSINESS_FAX_NUMBER);
                BusinessHomePage = GetMapiPropertyString(MapiTags.PR_BUSINESS_HOME_PAGE);

                // WorkAddress is only filled when the msg object is made with Outlook 2007 or higher.
                // So we fill it from code.
                if (string.IsNullOrEmpty(WorkAddress))
                {
                    var workAddress = string.Empty;

                    if (!string.IsNullOrEmpty(BusinessAddressStreet))
                        workAddress += BusinessAddressStreet + Environment.NewLine;

                    if (!string.IsNullOrEmpty(BusinessAddressPostalCode))
                        workAddress += BusinessAddressPostalCode + Environment.NewLine;

                    if (!string.IsNullOrEmpty(BusinessAddressCity))
                        workAddress += BusinessAddressCity + Environment.NewLine;

                    if (!string.IsNullOrEmpty(BusinessAddressCountry))
                        workAddress += BusinessAddressCountry + Environment.NewLine;

                    if (!string.IsNullOrEmpty(workAddress))
                        WorkAddress = workAddress;
                }
                #endregion

                #region Home address information
                HomeAddress = GetMapiPropertyString(MapiTags.HomeAddress);
                HomeAddressStreet = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_STREET);
                HomeAddressCity = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_CITY);
                HomeAddressState = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_STATE_OR_PROVINCE);
                HomeAddressPostalCode = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_POSTAL_CODE);
                HomeAddressCountry = GetMapiPropertyString(MapiTags.PR_HOME_ADDRESS_COUNTRY);
                HomeTelephoneNumber = GetMapiPropertyString(MapiTags.PR_HOME_TELEPHONE_NUMBER);
                HomeTelephoneNumber2 = GetMapiPropertyString(MapiTags.PR_HOME2_TELEPHONE_NUMBER);
                HomeFaxNumber = GetMapiPropertyString(MapiTags.PR_HOME_FAX_NUMBER);

                // HomeAddress is only filled when the msg object is made with Outlook 2007 or higher.
                // So we fill it from code.
                if (string.IsNullOrEmpty(HomeAddress))
                {
                    var homeAddress = string.Empty;

                    if (!string.IsNullOrEmpty(HomeAddressStreet))
                        homeAddress += HomeAddressStreet + Environment.NewLine;

                    if (!string.IsNullOrEmpty(HomeAddressPostalCode))
                        homeAddress += HomeAddressPostalCode + Environment.NewLine;

                    if (!string.IsNullOrEmpty(HomeAddressCity))
                        homeAddress += HomeAddressCity + Environment.NewLine;

                    if (!string.IsNullOrEmpty(HomeAddressCountry))
                        homeAddress += HomeAddressCountry + Environment.NewLine;

                    if (!string.IsNullOrEmpty(homeAddress))
                        HomeAddress = homeAddress;
                }
                #endregion

                #region Other address information
                OtherAddress = GetMapiPropertyString(MapiTags.OtherAddress);
                OtherAddressStreet = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_STREET);
                OtherAddressCity = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_CITY);
                OtherAddressState = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_STATE_OR_PROVINCE);
                OtherAddressPostalCode = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_POSTAL_CODE);
                OtherAddressCountry = GetMapiPropertyString(MapiTags.PR_OTHER_ADDRESS_COUNTRY);
                OtherTelephoneNumber = GetMapiPropertyString(MapiTags.PR_OTHER_TELEPHONE_NUMBER);

                // OtherAddress is only filled when the msg object is made with Outlook 2007 or higher.
                // So we fill it from code.
                if (string.IsNullOrEmpty(OtherAddress))
                {
                    var otherAddress = string.Empty;

                    if (!string.IsNullOrEmpty(OtherAddressStreet))
                        otherAddress += OtherAddressStreet + Environment.NewLine;

                    if (!string.IsNullOrEmpty(OtherAddressPostalCode))
                        otherAddress += OtherAddressPostalCode + Environment.NewLine; 
                    
                    if (!string.IsNullOrEmpty(OtherAddressCity))
                        otherAddress += OtherAddressCity + Environment.NewLine;

                    if (!string.IsNullOrEmpty(OtherAddressCountry))
                        otherAddress += OtherAddressCountry + Environment.NewLine;

                    if (!string.IsNullOrEmpty(otherAddress))
                        OtherAddress = otherAddress;
                }
                #endregion

                #region Primary information
                PrimaryTelephoneNumber = GetMapiPropertyString(MapiTags.PR_PRIMARY_TELEPHONE_NUMBER);
                PrimaryFaxNumber = GetMapiPropertyString(MapiTags.PR_PRIMARY_FAX_NUMBER);
                #endregion

                #region Assistant information
                AssistantName = GetMapiPropertyString(MapiTags.PR_ASSISTANT); 
                AssistantTelephoneNumber = GetMapiPropertyString(MapiTags.PR_ASSISTANT_TELEPHONE_NUMBER);
                #endregion

                InstantMessagingAddress = GetMapiPropertyString(MapiTags.InstantMessagingAddress);

                #region Telephone numbers
                CompanyMainTelephoneNumber = GetMapiPropertyString(MapiTags.PR_COMPANY_MAIN_PHONE_NUMBER);
                CellularTelephoneNumber = GetMapiPropertyString(MapiTags.PR_CELLULAR_TELEPHONE_NUMBER);
                CarTelephoneNumber = GetMapiPropertyString(MapiTags.PR_CAR_TELEPHONE_NUMBER);
                RadioTelephoneNumber = GetMapiPropertyString(MapiTags.PR_RADIO_TELEPHONE_NUMBER);
                BeeperTelephoneNumber = GetMapiPropertyString(MapiTags.PR_BEEPER_TELEPHONE_NUMBER);
                CallbackTelephoneNumber = GetMapiPropertyString(MapiTags.PR_CALLBACK_TELEPHONE_NUMBER);
                TextTelephone = GetMapiPropertyString(MapiTags.PR_TELEX_NUMBER);
                ISDNNumber = GetMapiPropertyString(MapiTags.PR_ISDN_NUMBER);
                TelexNumber = GetMapiPropertyString(MapiTags.PR_TELEX_NUMBER);
                #endregion

                #region E-mail information
                Email1EmailAddress = GetMapiPropertyString(MapiTags.Email1EmailAddress);
                Email1DisplayName = GetMapiPropertyString(MapiTags.Email1DisplayName);
                Email2EmailAddress = GetMapiPropertyString(MapiTags.Email2EmailAddress);
                Email2DisplayName = GetMapiPropertyString(MapiTags.Email2DisplayName);
                Email3EmailAddress = GetMapiPropertyString(MapiTags.Email3EmailAddress);
                Email3DisplayName = GetMapiPropertyString(MapiTags.Email3DisplayName);
                #endregion

                var birthday = GetMapiPropertyDateTime(MapiTags.PR_BIRTHDAY);
                if (birthday != null)
                    Birthday = ((DateTime) birthday).ToLocalTime();

                var weddingAnniversary = GetMapiPropertyDateTime(MapiTags.PR_WEDDING_ANNIVERSARY);
                if (weddingAnniversary != null)
                    WeddingAnniversary = ((DateTime) weddingAnniversary).ToLocalTime();

                SpouseName = GetMapiPropertyString(MapiTags.PR_SPOUSE_NAME);
                Profession = GetMapiPropertyString(MapiTags.PR_PROFESSION);
                Html = GetMapiPropertyString(MapiTags.Html);
            }
            #endregion
        }
    }
}
