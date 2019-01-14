//
// Appointment.cs
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
using System.Collections.Generic;
using System.Collections.ObjectModel;
using MsgReader.Localization;

namespace MsgReader.Outlook
{
    #region Enum AppointmentRecurrenceType
    /// <summary>
    ///     The recurrence type of an appointment
    /// </summary>
    public enum AppointmentRecurrenceType
    {
        /// <summary>
        ///     There is no reccurence
        /// </summary>
        None = -1,

        /// <summary>
        ///     The appointment is daily
        /// </summary>
        Daily = 0,

        /// <summary>
        ///     The appointment is weekly
        /// </summary>
        Weekly = 1,

        /// <summary>
        ///     The appointment is monthly
        /// </summary>
        Montly = 2,

        /// <summary>
        ///     The appointment is yearly
        /// </summary>
        Yearly = 3
    }
    #endregion

    #region Enum AppointmentClientIntent
    /// <summary>
    ///     The intent of an appointment
    /// </summary>
    public enum AppointmentClientIntent
    {
        /// <summary>
        ///     The user is the owner of the Meeting object's
        /// </summary>
        Manager = 1,

        /// <summary>
        ///     The user is a delegate acting on a Meeting object in a delegator's Calendar folder. If this bit is set, the
        ///     ciManager bit SHOULD NOT be set
        /// </summary>
        Delegate = 2,

        /// <summary>
        ///     The user deleted the Meeting object with no response sent to the organizer
        /// </summary>
        DeletedWithNoResponse = 4,

        /// <summary>
        ///     The user deleted an exception to a recurring series with no response sent to the organizer
        /// </summary>
        DeletedExceptionWithNoResponse = 8,

        /// <summary>
        ///     Appointment accepted as tentative
        /// </summary>
        RespondedTentative = 16,

        /// <summary>
        ///     Appointment accepted
        /// </summary>
        RespondedAccept = 32,

        /// <summary>
        ///     Appointment declined
        /// </summary>
        RespondedDecline = 64,

        /// <summary>
        ///     The user modified the start time
        /// </summary>
        ModifiedStartTime = 128,

        /// <summary>
        ///     The user modified the end time
        /// </summary>
        ModifiedEndTime = 256,

        /// <summary>
        ///     The user changed the location of the meeting
        /// </summary>
        ModifiedLocation = 512,

        /// <summary>
        ///     The user declined an exception to a recurring series
        /// </summary>
        RespondedExceptionDecline = 1024,

        /// <summary>
        ///     The user declined an exception to a recurring series
        /// </summary>
        Canceled = 2048,

        /// <summary>
        ///     The user canceled an exception to a recurring serie
        /// </summary>
        ExceptionCanceled = 4096
    }
    #endregion

    public partial class Storage
    {
        /// <summary>
        ///     Class used to contain all the appointment information of a <see cref="Storage.Message" />.
        /// </summary>
        public sealed class Appointment : Storage
        {
            #region Properties
            /// <summary>
            ///     Returns the location for the appointment, null when not available
            /// </summary>
            public string Location { get; }

            /// <summary>
            ///     Returns the start time for the appointment, null when not available
            /// </summary>
            public DateTime? Start { get; }

            /// <summary>
            ///     Returns the end time for the appointment, null when not available
            /// </summary>
            public DateTime? End { get; }

            /// <summary>
            ///     Returns a string with all the attendees (To and CC), if you also want their E-mail addresses then
            ///     get the <see cref="Storage.Message.Recipients" /> from the message, null when not available
            /// </summary>
            public string AllAttendees { get; }

            /// <summary>
            ///     Returns a string with all the TO (mandatory) attendees. If you also want their E-mail addresses then
            ///     get the <see cref="Storage.Message.Recipients" /> from the <see cref="Storage.Message" /> and filter this
            ///     one on <see cref="RecipientType.To" />. Null when not available
            /// </summary>
            public string ToAttendees { get; }

            /// <summary>
            ///     Returns a string with all the CC (optional) attendees. If you also want their E-mail addresses then
            ///     get the <see cref="Storage.Message.Recipients" /> from the <see cref="Storage.Message" /> and filter this
            ///     one on <see cref="RecipientType.Cc" />. Null when not available
            /// </summary>
            public string CcAttendees { get; }

            /// <summary>
            ///     Returns A value of <c>true</c> for the PidLidAppointmentNotAllowPropose property ([MS-OXPROPS] section 2.17)
            ///     indicates that attendees are not allowed to propose a new date and/or time for the meeting. A value of
            ///     <c>false</c> or the absence of this property indicates that the attendees are allowed to propose a new date
            ///     and/or time. This property is meaningful only on Meeting objects, Meeting Request objects, and Meeting
            ///     Update objects. Null when not available
            /// </summary>
            public bool? NotAllowPropose { get; }

            /// <summary>
            ///     Returns a <see cref="UnsendableRecipients" /> object with all the unsendable attendees. Null when not available
            /// </summary>
            public UnsendableRecipients UnsendableRecipients { get; }

            /// <summary>
            ///     Returns the reccurence type (daily, weekly, monthly or yearly) for the <see cref="Storage.Appointment" />
            /// </summary>
            public AppointmentRecurrenceType ReccurrenceType { get; }

            /// <summary>
            ///     Returns the reccurence type (daily, weekly, monthly or yearly) for the <see cref="Storage.Appointment" /> as a
            ///     string,
            ///     null when not available
            /// </summary>
            public string RecurrenceTypeText { get; }

            /// <summary>
            ///     Returns the reccurence patern for the <see cref="Storage.Appointment" />, null when not available
            /// </summary>
            public string RecurrencePatern { get; }

            /// <summary>
            ///     The clients intention for the the <see cref="Storage.Appointment" /> as a list,
            ///     null when not available
            ///     of <see cref="AppointmentClientIntent" />
            /// </summary>
            public ReadOnlyCollection<AppointmentClientIntent> ClientIntent { get; }

            /// <summary>
            ///     The <see cref="ClientIntent" /> for the the <see cref="Storage.Appointment" /> as text
            /// </summary>
            public string ClientIntentText { get; }
            #endregion

            #region Constructor
            /// <summary>
            ///     Initializes a new instance of the <see cref="Storage.Task" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Appointment(Storage message) : base(message._rootStorage)
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
                {
                    ClientIntent = null;
                }
                else
                {
                    var bitwiseValue = (int) clientIntent;

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