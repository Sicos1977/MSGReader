using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using MsgReader.Localization;

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
        /// Class used to contain all the appointment information of a <see cref="Storage.Message"/>.
        /// </summary>
        public sealed class Appointment : Storage
        {
            #region Public enum AppointmentRecurrenceType
            /// <summary>
            /// The recurrence type of an appointment
            /// </summary>
            public enum AppointmentRecurrenceType
            {
                /// <summary>
                /// There is no reccurence
                /// </summary>
                None = -1,

                /// <summary>
                /// The appointment is daily
                /// </summary>
                Daily = 0,

                /// <summary>
                /// The appointment is weekly
                /// </summary>
                Weekly = 1,

                /// <summary>
                /// The appointment is monthly
                /// </summary>
                Montly = 2,

                /// <summary>
                /// The appointment is yearly
                /// </summary>
                Yearly = 3
            }
            #endregion

            #region Public enum AppointmentClientIntent
            /// <summary>
            /// The intent of an appointment
            /// </summary>
            public enum AppointmentClientIntent
            {
                /// <summary>
                /// The user is the owner of the Meeting object's
                /// </summary>
                Manager = 1,

                /// <summary>
                /// The user is a delegate acting on a Meeting object in a delegator's Calendar folder. If this bit is set, the ciManager bit SHOULD NOT be set
                /// </summary>
                Delegate = 2,

                /// <summary>
                /// The user deleted the Meeting object with no response sent to the organizer
                /// </summary>
                DeletedWithNoResponse = 4,

                /// <summary>
                /// The user deleted an exception to a recurring series with no response sent to the organizer
                /// </summary>
                DeletedExceptionWithNoResponse = 8,

                /// <summary>
                /// Appointment accepted as tentative
                /// </summary>
                RespondedTentative = 16,

                /// <summary>
                /// Appointment accepted
                /// </summary>
                RespondedAccept = 32,

                /// <summary>
                /// Appointment declined
                /// </summary>
                RespondedDecline = 64,

                /// <summary>
                /// The user modified the start time
                /// </summary>
                ModifiedStartTime = 128,

                /// <summary>
                /// The user modified the end time
                /// </summary>
                ModifiedEndTime = 256,

                /// <summary>
                /// The user changed the location of the meeting
                /// </summary>
                ModifiedLocation = 512,

                /// <summary>
                /// The user declined an exception to a recurring series
                /// </summary>
	            RespondedExceptionDecline = 1024,

                /// <summary>
                /// The user declined an exception to a recurring series
                /// </summary>
	            Canceled = 2048,

                /// <summary>
                /// The user canceled an exception to a recurring serie
                /// </summary>
	            ExceptionCanceled = 4096
            }
            #endregion

            #region Properties
            /// <summary>
            /// Returns the location for the appointment, null when not available
            /// </summary>
            public string Location { get; private set; }

            /// <summary>
            /// Returns the start time for the appointment, null when not available
            /// </summary>
            public DateTime? Start { get; private set; }
            
            /// <summary>
            /// Returns the end time for the appointment, null when not available
            /// </summary>
            public DateTime? End { get; private set; }

            /// <summary>
            /// Returns a string with all the attendees (To and CC), if you also want their E-mail addresses then
            /// get the <see cref="Storage.Message.Recipients"/> from the message, null when not available
            /// </summary>
            public string AllAttendees { get; private set; }

            /// <summary>
            /// Returns a string with all the TO (mandatory) attendees. If you also want their E-mail addresses then
            /// get the <see cref="Storage.Message.Recipients"/> from the <see cref="Storage.Message"/> and filter this 
            /// one on <see cref="Storage.Recipient.RecipientType.To"/>. Null when not available
            /// </summary>
            public string ToAttendees { get; private set; }

            /// <summary>
            /// Returns a string with all the CC (optional) attendees. If you also want their E-mail addresses then
            /// get the <see cref="Storage.Message.Recipients"/> from the <see cref="Storage.Message"/> and filter this 
            /// one on <see cref="Storage.Recipient.RecipientType.Cc"/>. Null when not available
            /// </summary>
            public string CcAttendees { get; private set; }

            /// <summary>
            /// Returns A value of <c>true</c> for the PidLidAppointmentNotAllowPropose property ([MS-OXPROPS] section 2.17) 
            /// indicates that attendees are not allowed to propose a new date and/or time for the meeting. A value of 
            /// <c>false</c> or the absence of this property indicates that the attendees are allowed to propose a new date 
            /// and/or time. This property is meaningful only on Meeting objects, Meeting Request objects, and Meeting 
            /// Update objects. Null when not available
            /// </summary>
            public bool? NotAllowPropose { get; private set; }

            /// <summary>
            /// Returns a <see cref="UnsendableRecipients"/> object with all the unsendable attendees. Null when not available
            /// </summary>
            public UnsendableRecipients UnsendableRecipients { get; private set; }

            /// <summary>
            /// Returns the reccurence type (daily, weekly, monthly or yearly) for the <see cref="Storage.Appointment"/>
            /// </summary>
            public AppointmentRecurrenceType ReccurrenceType { get; private set; }

            /// <summary>
            /// Returns the reccurence type (daily, weekly, monthly or yearly) for the <see cref="Storage.Appointment"/> as a string, 
            /// null when not available
            /// </summary>
            public string RecurrenceTypeText { get; private set; }

            /// <summary>
            /// Returns the reccurence patern for the <see cref="Storage.Appointment"/>, null when not available
            /// </summary>
            public string RecurrencePatern { get; private set; }

            /// <summary>
            /// The clients intention for the the <see cref="Storage.Appointment"/> as a list,
            /// null when not available
            /// of <see cref="AppointmentClientIntent"/>
            /// </summary>
            public ReadOnlyCollection<AppointmentClientIntent> ClientIntent { get; private set; }

            /// <summary>
            /// The <see cref="ClientIntent"/> for the the <see cref="Storage.Appointment"/> as text
            /// </summary>
            public string ClientIntentText { get; private set; }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Task" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Appointment(Storage message) : base(message._storage)
            {
                //GC.SuppressFinalize(message);
                _namedProperties = message._namedProperties;
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
               
                Location = GetMapiPropertyString(MapiTags.Location);
                Start = GetMapiPropertyDateTime(MapiTags.AppointmentStartWhole);
                End = GetMapiPropertyDateTime(MapiTags.AppointmentEndWhole);
                AllAttendees = GetMapiPropertyString(MapiTags.AppointmentAllAttendees);
                ToAttendees = GetMapiPropertyString(MapiTags.AppointmentToAttendees);
                CcAttendees = GetMapiPropertyString(MapiTags.AppointmentCCAttendees);
                NotAllowPropose = GetMapiPropertyBool(MapiTags.AppointmentNotAllowPropose);
                UnsendableRecipients = GetUnsendableRecipients(MapiTags.AppointmentUnsendableRecipients);

                #region Recurrence
                var recurrenceType = GetMapiPropertyInt32(MapiTags.ReccurrenceType);
                if (recurrenceType == null)
                {
                    ReccurrenceType = AppointmentRecurrenceType.None;
                    RecurrenceTypeText = LanguageConsts.AppointmentReccurenceTypeNoneText;
                }
                else
                {
                    switch (recurrenceType)
                    {
                        case 1:
                            ReccurrenceType = AppointmentRecurrenceType.Daily;
                            break;

                        case 2:
                            ReccurrenceType = AppointmentRecurrenceType.Weekly;
                            break;

                        case 3:
                        case 4:
                            ReccurrenceType = AppointmentRecurrenceType.Montly;
                            break;

                        case 5:
                        case 6:
                            ReccurrenceType = AppointmentRecurrenceType.Yearly;
                            break;

                        default:
                            ReccurrenceType = AppointmentRecurrenceType.None;
                            break;
                    }
                
                    switch (ReccurrenceType)
                    {
                        case AppointmentRecurrenceType.Daily:
                            RecurrenceTypeText = LanguageConsts.AppointmentReccurenceTypeDailyText;
                            break;

                        case AppointmentRecurrenceType.Weekly:
                            RecurrenceTypeText = LanguageConsts.AppointmentReccurenceTypeWeeklyText;
                            break;

                        case AppointmentRecurrenceType.Montly:
                            RecurrenceTypeText = LanguageConsts.AppointmentReccurenceTypeMonthlyText;
                            break;

                        case AppointmentRecurrenceType.Yearly:
                            RecurrenceTypeText = LanguageConsts.AppointmentReccurenceTypeYearlyText;
                            break;
                    }
                }

                RecurrencePatern = GetMapiPropertyString(MapiTags.ReccurrencePattern);
                #endregion

                #region ClientIntent
                var clientIntentList = new List<AppointmentClientIntent>();
                var clientIntent = GetMapiPropertyInt32(MapiTags.PidLidClientIntent);

                if (clientIntent == null)
                    ClientIntent = null;
                else
                {
                    var bitwiseValue = (int)clientIntent;

                    if ((bitwiseValue & 1) == 1)
                    {
                        clientIntentList.Add(AppointmentClientIntent.Manager);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentManagerText;
                    }

                    if ((bitwiseValue & 2) == 2)
                    {
                        clientIntentList.Add(AppointmentClientIntent.Delegate);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentDelegateText;
                    }

                    if ((bitwiseValue & 4) == 4)
                    {
                        clientIntentList.Add(AppointmentClientIntent.DeletedWithNoResponse);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentDeletedWithNoResponseText;
                    }

                    if ((bitwiseValue & 8) == 8)
                    {
                        clientIntentList.Add(AppointmentClientIntent.DeletedExceptionWithNoResponse);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentDeletedExceptionWithNoResponseText;
                    }

                    if ((bitwiseValue & 16) == 16)
                    {
                        clientIntentList.Add(AppointmentClientIntent.RespondedTentative);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentRespondedTentativeText;
                    }

                    if ((bitwiseValue & 32) == 32)
                    {
                        clientIntentList.Add(AppointmentClientIntent.RespondedAccept);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentRespondedAcceptText;
                    }

                    if ((bitwiseValue & 64) == 64)
                    {
                        clientIntentList.Add(AppointmentClientIntent.RespondedDecline);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentRespondedDeclineText;
                    }
                    if ((bitwiseValue & 128) == 128)
                    {
                        clientIntentList.Add(AppointmentClientIntent.ModifiedStartTime);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentModifiedStartTimeText;
                    }

                    if ((bitwiseValue & 256) == 256)
                    {
                        clientIntentList.Add(AppointmentClientIntent.ModifiedEndTime);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentModifiedEndTimeText;
                    }

                    if ((bitwiseValue & 512) == 512)
                    {
                        clientIntentList.Add(AppointmentClientIntent.ModifiedLocation);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentModifiedLocationText;
                    }

                    if ((bitwiseValue & 1024) == 1024)
                    {
                        clientIntentList.Add(AppointmentClientIntent.RespondedExceptionDecline);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentRespondedExceptionDeclineText;
                    }

                    if ((bitwiseValue & 2048) == 2048)
                    {
                        clientIntentList.Add(AppointmentClientIntent.Canceled);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentCanceledText;
                    }

                    if ((bitwiseValue & 4096) == 4096)
                    {
                        clientIntentList.Add(AppointmentClientIntent.ExceptionCanceled);
                        ClientIntentText = LanguageConsts.AppointmentClientIntentExceptionCanceledText;
                    }

                    ClientIntent = clientIntentList.AsReadOnly();
                }
                #endregion
            }
            #endregion
        }
    }
}