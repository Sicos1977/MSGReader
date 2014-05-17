using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    public partial class Storage
    {
        /// <summary>
        /// Class used to contain all the appointment information of a <see cref="Storage.Message"/>.
        /// </summary>
        public sealed class Appointment : Storage
        {
            #region Public enum AppointmentRecurrenceType
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
            /// Returns the reccurence type (daily, weekly, monthly or yearly) for the appointment
            /// </summary>
            public AppointmentRecurrenceType ReccurrenceType { get; private set; }

            /// <summary>
            /// Returns the reccurence type (daily, weekly, monthly or yearly) for the appointment as a string
            /// </summary>
            public string RecurrenceTypeText
            {
                get
                {
                    switch (ReccurrenceType)
                    {
                        case AppointmentRecurrenceType.Daily:
                            return LanguageConsts.AppointmentReccurenceTypeDailyText;

                        case AppointmentRecurrenceType.Weekly:
                            return LanguageConsts.AppointmentReccurenceTypeWeeklyText;

                        case AppointmentRecurrenceType.Montly:
                            return LanguageConsts.AppointmentReccurenceTypeMonthlyText;

                        case AppointmentRecurrenceType.Yearly:
                            return LanguageConsts.AppointmentReccurenceTypeYearlyText;
                    }

                    return LanguageConsts.AppointmentReccurenceTypeNoneText;
                }
            }

            /// <summary>
            /// Returns the reccurence patern for the appointment, null when not available
            /// </summary>
            public string RecurrencePatern { get; private set; }

            /// <summary>
            /// The clients intention for the the <see cref="Storage.Appointment"/> as a list 
            /// of <see cref="AppointmentClientIntent"/>
            /// </summary>
            public ReadOnlyCollection<AppointmentClientIntent> ClientIntent { get; private set; }

            /// <summary>
            /// The clients intention for the the <see cref="Storage.Appointment"/> as text
            /// </summary>
            public string ClientIntentText
            {
                get
                {
                    var clientIntent = ClientIntent;
                    if (clientIntent == null)
                        return null;

                    if (clientIntent.Contains(AppointmentClientIntent.Manager))
                        return LanguageConsts.AppointmentClientIntentManagerText;

                    if (clientIntent.Contains(AppointmentClientIntent.Manager))
                        return LanguageConsts.AppointmentClientIntentDelegateText;

                    if (ClientIntent.Contains(AppointmentClientIntent.DeletedWithNoResponse))
                        return LanguageConsts.AppointmentClientIntentDeletedWithNoResponseText;
                    
                    if (ClientIntent.Contains(AppointmentClientIntent.DeletedExceptionWithNoResponse))
                        return LanguageConsts.AppointmentClientIntentDeletedExceptionWithNoResponseText;

                    if (ClientIntent.Contains(AppointmentClientIntent.RespondedTentative))
                        return LanguageConsts.AppointmentClientIntentRespondedTentativeText;

                    if (ClientIntent.Contains(AppointmentClientIntent.RespondedAccept))
                        return LanguageConsts.AppointmentClientIntentRespondedAcceptText;

                    if (ClientIntent.Contains(AppointmentClientIntent.RespondedDecline))
                        return LanguageConsts.AppointmentClientIntentRespondedDeclineText;

                    if (ClientIntent.Contains(AppointmentClientIntent.ModifiedStartTime))
                        return LanguageConsts.AppointmentClientIntentModifiedStartTimeText;

                    if (ClientIntent.Contains(AppointmentClientIntent.ModifiedEndTime))
                        return LanguageConsts.AppointmentClientIntentModifiedEndTimeText;

                    if (ClientIntent.Contains(AppointmentClientIntent.ModifiedLocation))
                        return LanguageConsts.AppointmentClientIntentModifiedLocationText;

                    if (ClientIntent.Contains(AppointmentClientIntent.RespondedExceptionDecline))
                        return LanguageConsts.AppointmentClientIntentRespondedExceptionDeclineText;

                    if (ClientIntent.Contains(AppointmentClientIntent.Canceled))
                        return LanguageConsts.AppointmentClientIntentCanceledText;
                    
                    if (ClientIntent.Contains(AppointmentClientIntent.ExceptionCanceled))
                        return LanguageConsts.AppointmentClientIntentExceptionCanceledText; 
                    
                    return null;
                }
            }
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

                var recurrenceType = GetMapiPropertyInt32(MapiTags.ReccurrenceType);
                if (recurrenceType == null)
                {
                    ReccurrenceType = AppointmentRecurrenceType.None;
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
                }

                RecurrencePatern = GetMapiPropertyString(MapiTags.ReccurrencePattern);

                // ClientIntent
                var clientIntentList = new List<AppointmentClientIntent>();


                var clientIntent = GetMapiPropertyInt32(MapiTags.PidLidClientIntent);

                if (clientIntent == null)
                    ClientIntent = null;
                else
                {
                    var bitwiseValue = (int)clientIntent;

                    if ((bitwiseValue & 1) == 1)
                        clientIntentList.Add(AppointmentClientIntent.Manager);

                    if ((bitwiseValue & 2) == 2)
                        clientIntentList.Add(AppointmentClientIntent.Delegate);

                    if ((bitwiseValue & 4) == 4)
                        clientIntentList.Add(AppointmentClientIntent.DeletedWithNoResponse);

                    if ((bitwiseValue & 8) == 8)
                        clientIntentList.Add(AppointmentClientIntent.DeletedExceptionWithNoResponse);

                    if ((bitwiseValue & 16) == 16)
                        clientIntentList.Add(AppointmentClientIntent.RespondedTentative);

                    if ((bitwiseValue & 32) == 32)
                        clientIntentList.Add(AppointmentClientIntent.RespondedAccept);

                    if ((bitwiseValue & 64) == 64)
                        clientIntentList.Add(AppointmentClientIntent.RespondedDecline);

                    if ((bitwiseValue & 128) == 128)
                        clientIntentList.Add(AppointmentClientIntent.ModifiedStartTime);

                    if ((bitwiseValue & 256) == 256)
                        clientIntentList.Add(AppointmentClientIntent.ModifiedEndTime);

                    if ((bitwiseValue & 512) == 512)
                        clientIntentList.Add(AppointmentClientIntent.ModifiedLocation);

                    if ((bitwiseValue & 1024) == 1024)
                        clientIntentList.Add(AppointmentClientIntent.RespondedExceptionDecline);

                    if ((bitwiseValue & 2048) == 2048)
                        clientIntentList.Add(AppointmentClientIntent.Canceled);

                    if ((bitwiseValue & 4096) == 4096)
                        clientIntentList.Add(AppointmentClientIntent.ExceptionCanceled);

                    ClientIntent = clientIntentList.AsReadOnly();
                }
            }
            #endregion
        }
    }
}