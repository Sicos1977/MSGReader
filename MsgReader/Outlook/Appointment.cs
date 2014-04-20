using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal partial class Storage
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
            /// Returns the location for the appointment
            /// </summary>
            public string Location
            {
                get { return GetMapiPropertyString(MapiTags.Location); }
            }

            /// <summary>
            /// Returns the start time for the appointment
            /// </summary>
            public DateTime? Start
            {
                get { return GetMapiPropertyDateTime(MapiTags.AppointmentStartWhole); }
            }
            
            /// <summary>
            /// Returns the end time for the appointment
            /// </summary>
            public DateTime? End
            {
                get { return GetMapiPropertyDateTime(MapiTags.AppointmentEndWhole); }
            }

            /// <summary>
            /// Returns the reccurence type (daily, weekly, monthly or yearly) for the appointment
            /// </summary>
            public AppointmentRecurrenceType ReccurrenceType
            {
                get
                {
                    var value = GetMapiPropertyInt32(MapiTags.ReccurrenceType);
                    switch (value)
                    {
                        case 1:
                            return AppointmentRecurrenceType.Daily;

                        case 2:
                            return AppointmentRecurrenceType.Weekly;

                        case 3:
                        case 4:
                            return AppointmentRecurrenceType.Montly;

                        case 5:
                        case 6:
                            return AppointmentRecurrenceType.Yearly;
                    }

                    return AppointmentRecurrenceType.None;
                }
            }

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
            /// Returns the reccurence patern for the appointment
            /// </summary>
            public string RecurrencePatern
            {
                get { return GetMapiPropertyString(MapiTags.ReccurrencePattern); }
            }

            /// <summary>
            /// The clients intention to the appointment or null when unknown
            /// </summary>
            public ReadOnlyCollection<AppointmentClientIntent> ClientIntent
            {
                get
                {
                    // ClientIntent
                    var result = new List<AppointmentClientIntent>();

                    try
                    {
                        var value = GetMapiPropertyInt32(MapiTags.PidLidClientIntent);

                        if (value == null)
                            return null;

                        var bitwiseValue = (int) value;

                        if ((bitwiseValue & 1) == 1)
                            result.Add(AppointmentClientIntent.Manager);

                        if ((bitwiseValue & 2) == 2)
                            result.Add(AppointmentClientIntent.Delegate);

                        if ((bitwiseValue & 4) == 4)
                            result.Add(AppointmentClientIntent.DeletedWithNoResponse);

                        if ((bitwiseValue & 8) == 8)
                            result.Add(AppointmentClientIntent.DeletedExceptionWithNoResponse);

                        if ((bitwiseValue & 16) == 16)
                            result.Add(AppointmentClientIntent.RespondedTentative);

                        if ((bitwiseValue & 32) == 32)
                            result.Add(AppointmentClientIntent.RespondedAccept);

                        if ((bitwiseValue & 64) == 64)
                            result.Add(AppointmentClientIntent.RespondedDecline);

                        if ((bitwiseValue & 128) == 128)
                            result.Add(AppointmentClientIntent.ModifiedStartTime);

                        if ((bitwiseValue & 256) == 256)
                            result.Add(AppointmentClientIntent.ModifiedEndTime);

                        if ((bitwiseValue & 512) == 512)
                            result.Add(AppointmentClientIntent.ModifiedLocation);

                        if ((bitwiseValue & 1024) == 1024)
                            result.Add(AppointmentClientIntent.RespondedExceptionDecline);

                        if ((bitwiseValue & 2048) == 2048)
                            result.Add(AppointmentClientIntent.Canceled);

                        if ((bitwiseValue & 4096) == 4096)
                            result.Add(AppointmentClientIntent.ExceptionCanceled);
                        return result.AsReadOnly();
                    }
                    catch (NullReferenceException)
                    {
                        return null;
                    }
                }
            }

            /// <summary>
            /// The clients intention to the appointment as text
            /// </summary>
            public string ClientIntentText
            {
                get
                {
                    var status = ClientIntent;
                    if (status == null)
                        return null;

                    if (status.Contains(AppointmentClientIntent.Manager))
                        return LanguageConsts.AppointmentClientIntentManagerText;

                    if (status.Contains(AppointmentClientIntent.Manager))
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
            internal Appointment(Storage message)
                : base(message._storage)
            {
                GC.SuppressFinalize(message);
                _namedProperties = message._namedProperties;
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
            }
            #endregion
        }
    }
}