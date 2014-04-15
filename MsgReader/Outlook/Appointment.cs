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
            public string RecurrenceType
            {
                get
                {
                    var value = GetMapiPropertyInt32(MapiTags.ReccurrenceType);
                    switch (value)
                    {
                        case 1:
                            return LanguageConsts.AppointmentReccurenceTypeDailyText;

                        case 2:
                            return LanguageConsts.AppointmentReccurenceTypeWeeklyText;

                        case 3:
                        case 4:
                            return LanguageConsts.AppointmentReccurenceTypeMonthlyText;

                        case 5:
                        case 6:
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
            /// The appointment status, or null when unkown
            /// </summary>
            public string Status
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
                                return LanguageConsts.AppointmentStatusManager;

                            case 16:
                                return LanguageConsts.AppointmentStatusTentative;

                            case 32:
                                return LanguageConsts.AppointmentStatusAccept;

                            case 64:
                                return LanguageConsts.AppointmentStatusDecline;

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