using System;

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

            #region Public enum AppointmentStatus
            public enum AppointmentStatus
            {
                /// <summary>
                /// The manager of the appointment
                /// </summary>
                Manager = 1,

                /// <summary>
                /// Appointment accepted as tentative
                /// </summary>
                Tentative = 16,

                /// <summary>
                /// Appointment accepted
                /// </summary>
                Accept = 32,

                /// <summary>
                /// Appointment declined
                /// </summary>
                Decline = 64
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

                    return LanguageConsts.AppointmentReccurenceTypeNone;
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
            /// The appointment status or null when unknown
            /// </summary>
            public AppointmentStatus? Status
            {
                get
                {
                    // ClientIntent
                    try
                    {
                        var value = GetMapiPropertyInt32(MapiTags.ClientIntent);
                        switch (value)
                        {
                            case 1:
                                return AppointmentStatus.Manager;

                            case 16:
                                return AppointmentStatus.Tentative;

                            case 32:
                                return AppointmentStatus.Accept;

                            case 64:
                                return AppointmentStatus.Decline;

                            default:
                                return null;
                        }
                    }
                    catch (NullReferenceException)
                    {
                        return null;
                    }
                }
            }

            /// <summary>
            /// The appointment status as text or null when unknown
            /// </summary>
            public string StatusText
            {
                get
                {
                    // ClientIntent
                    switch (Status)
                    {
                        case AppointmentStatus.Manager:
                            return LanguageConsts.AppointmentStatusManager;

                        case AppointmentStatus.Tentative:
                            return LanguageConsts.AppointmentStatusTentative;

                        case AppointmentStatus.Accept:
                            return LanguageConsts.AppointmentStatusAccept;

                        case AppointmentStatus.Decline:
                            return LanguageConsts.AppointmentStatusDecline;

                        default:
                            return null;
                    }
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
                _namedProperties = message._namedProperties;
                GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
            }
            #endregion
        }
    }
}