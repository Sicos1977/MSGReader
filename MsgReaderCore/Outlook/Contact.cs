//
// Contact.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2018 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;

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
            public string DisplayName { get; }

            /// <summary>
            /// Returns the prefix of the name (e.g. De heer), null when not available
            /// </summary>
            public string Prefix { get; }

            /// <summary>
            /// Returns the initials, null when not available
            /// </summary>
            public string Initials { get; }

            /// <summary>
            /// Returns the sur name (e.g. Spelde), null when not available
            /// </summary>
            public string SurName { get; }

            /// <summary>
            /// Returns the given name (e.g. Kees), null when not available
            /// </summary>
            public string GivenName { get; }

            /// <summary>
            /// Returns the generation (e.g. Jr.), null when not available
            /// </summary>
            public string Generation { get; }

            /// <summary>
            /// Returns the function, null when not available
            /// </summary>
            public string Function { get; }

            /// <summary>
            /// Returns the department, null when not available
            /// </summary>
            public string Department { get; }

            /// <summary>
            /// Returns the name of the company, null when not available
            /// </summary>
            public string Company { get; }

            #region Business address information
            /// <summary>
            /// Returns the named propery work (business) address (Outlook 2007 or higher), null when not available
            /// </summary>
            public string WorkAddress { get; }

            /// <summary>
            /// Returns the street of the business address, null when not available
            /// </summary>
            public string BusinessAddressStreet { get; }

            /// <summary>
            /// Returns the city of the business address, null when not available
            /// </summary>
            public string BusinessAddressCity { get; }

            /// <summary>
            /// Returns the state of the business address, null when not available
            /// </summary>
            public string BusinessAddressState { get; }

            /// <summary>
            /// Returns the postal code of the business address, null when not available
            /// </summary>
            public string BusinessAddressPostalCode { get; }

            /// <summary>
            /// Returns the country of the business address, null when not available
            /// </summary>
            public string BusinessAddressCountry { get; }

            /// <summary>
            /// Returns the business telephone number, null when not available
            /// </summary>
            public string BusinessTelephoneNumber { get; }

            /// <summary>
            /// Returns the business second telephone number, null when not available
            /// </summary>
            public string BusinessTelephoneNumber2 { get; }
            
            /// <summary>
            /// Returns the business fax number, null when not available
            /// </summary>
            public string BusinessFaxNumber { get; }

            /// <summary>
            /// Returns the business home page, null when not available
            /// </summary>
            public string BusinessHomePage { get; }
            #endregion

            #region Home address information
            /// <summary>
            /// Returns the named propery home address (Outlook 2007 or higher)
            /// </summary>
            public string HomeAddress { get; }

            /// <summary>
            /// Returns the street of the home address, null when not available
            /// </summary>
            public string HomeAddressStreet { get; }

            /// <summary>
            /// Returns the city of the home address, null when not available
            /// </summary>
            public string HomeAddressCity { get; }

            /// <summary>
            /// Returns the state of the home address, null when not available
            /// </summary>
            public string HomeAddressState { get; }

            /// <summary>
            /// Returns the postal code of the home address, null when not available
            /// </summary>
            public string HomeAddressPostalCode { get; }

            /// <summary>
            /// Returns the country of the home address, null when not available
            /// </summary>
            public string HomeAddressCountry { get; }

            /// <summary>
            /// Returns the home telephone number, null when not available
            /// </summary>
            public string HomeTelephoneNumber { get; }

            /// <summary>
            /// Returns the home second telephone number, null when not available
            /// </summary>
            public string HomeTelephoneNumber2 { get; }

            /// <summary>
            /// Returns the home fax number, null when not available
            /// </summary>
            public string HomeFaxNumber { get; }
            #endregion

            #region Other address information
            /// <summary>
            /// Returns the named propery other address (Outlook 2007 or higher)
            /// </summary>
            public string OtherAddress { get; }

            /// <summary>
            /// Returns the street of the other address, null when not available
            /// </summary>
            public string OtherAddressStreet { get; }

            /// <summary>
            /// Returns the city of the other address, null when not available
            /// </summary>
            public string OtherAddressCity { get; }

            /// <summary>
            /// Returns the state of the other address, null when not available
            /// </summary>
            public string OtherAddressState { get; }

            /// <summary>
            /// Returns the postal code of the other address, null when not available
            /// </summary>
            public string OtherAddressPostalCode { get; }

            /// <summary>
            /// Returns the country of the other address, null when not available
            /// </summary>
            public string OtherAddressCountry { get; }

            /// <summary>
            /// Returns the other telephone number, null when not available
            /// </summary>
            public string OtherTelephoneNumber { get; }
            #endregion

            #region Primary information
            /// <summary>
            /// Returns the primary telephone number, null when not available
            /// </summary>
            public string PrimaryTelephoneNumber { get; }

            /// <summary>
            /// Returns the primary fax number, null when not available
            /// </summary>
            public string PrimaryFaxNumber { get; }
            #endregion

            #region Assistant information
            /// <summary>
            /// Returns the name of the assistant, null when not available
            /// </summary>
            public string AssistantName { get; }

            /// <summary>
            /// Returns the telephone number of the assistant, null when not available
            /// </summary>
            public string AssistantTelephoneNumber { get; }
            #endregion
            
            /// <summary>
            /// Return the instant messaging address, null when not available
            /// </summary>
            public string InstantMessagingAddress { get; }

            #region Telephone numbers
            /// <summary>
            /// Returns the company main telephone number, null when not available
            /// </summary>
            public string CompanyMainTelephoneNumber { get; }
            
            /// <summary>
            /// Returns the cellular telephone number, null when not available
            /// </summary>
            public string CellularTelephoneNumber { get; }
            
            /// <summary>
            /// Returns the car telephone number, null when not available
            /// </summary>
            public string CarTelephoneNumber { get; }
            
            /// <summary>
            /// Returns the radio telephone number, null when not available
            /// </summary>
            public string RadioTelephoneNumber { get; }

            /// <summary>
            /// Returns the beeper telephone number, null when not available
            /// </summary>
            public string BeeperTelephoneNumber { get; }

            /// <summary>
            /// Returns the callback telephone number, null when not available
            /// </summary>
            public string CallbackTelephoneNumber { get; }

            /// <summary>
            /// Returns the telex number, null when not available
            /// </summary>
            public string TelexNumber { get; }

            /// <summary>
            /// Returns the text telephone (TTYTDD), null when not available
            /// </summary>
            public string TextTelephone { get; }

            /// <summary>
            /// Returns the ISDN number, null when not available
            /// </summary>
            // ReSharper disable once InconsistentNaming
            public string ISDNNumber { get; }
            #endregion

            #region E-mail information
            /// <summary>
            /// Returns the name property e-mail address 1 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email1EmailAddress { get; }

            /// <summary>
            /// Returns the name property e-mail displayname 1 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email1DisplayName { get; }

            /// <summary>
            /// Returns the name property e-mail address 2 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email2EmailAddress { get; }

            /// <summary>
            /// Returns the name property e-mail displayname 2 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email2DisplayName { get; }
            
            /// <summary>
            /// Returns the name property e-mail address 3 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email3EmailAddress { get; }

            /// <summary>
            /// Returns the name property e-mail displayname 3 (Outlook 2007 or higher), null when not available
            /// </summary>
            public string Email3DisplayName { get; }
            #endregion

            /// <summary>
            /// Returns the birthday, null when not available
            /// </summary>
            public DateTime? Birthday { get; }

            /// <summary>
            /// Returns the wedding/anniversary, null when not available
            /// </summary>
            public DateTime? WeddingAnniversary { get; }

            /// <summary>
            /// Returns the name of the spouse, null when not available
            /// </summary>
            public string SpouseName { get; }

            /// <summary>
            /// Returns the profession, null when not available
            /// </summary>
            public string Profession { get; }

            /// <summary>
            /// Returns the homepage (html), null when not available
            /// </summary>
            public string Html { get; }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Contact" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Contact(Storage message) : base(message._rootStorage)
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
