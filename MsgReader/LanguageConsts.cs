namespace DocumentServices.Modules.Readers.MsgReader
{
    /// <summary>
    /// This class contains all the language specific constants
    /// </summary>
    internal static class LanguageConsts
    {
         /***************************************************************************************
         *  A note to you, please send me this file when you translate it to your own language. *
         *  You can send the file to : sicos2002@hotmail.com                                    *
         ***************************************************************************************/

        #region E-mail constants
        public const string EmailFromLabel = "From";
        public const string EmailSentOnLabel = "Sent on";
        public const string EmailToLabel = "To";
        public const string EmailCcLabel = "CC";
        public const string EmailBccLabel = "BCC";
        public const string EmailSubjectLabel = "Subject";
        public const string EmailAttachmentsLabel = "Attachments";
        public const string EmailFollowUpFlag = "Flag";
        public const string EmailFollowUpLabel = "Follow up";
        public const string EmailFollowUpStatusLabel = "Follow up status";
        public const string EmailFollowUpCompletedText = "Completed";
        public const string EmailCategoriesLabel = "Categories";
        #endregion

        /// <summary>
        /// The culture format to use for a date
        /// </summary>
        public const string DateFormatCulture = "en-US";

        /// <summary>
        /// The format used for all date/time related items, use this in conjuction with the <see cref="DateFormatCulture"/>
        /// http://msdn.microsoft.com/en-us/library/az4se3k1%28v=vs.110%29.aspx
        /// </summary>
        public const string DataFormatWithTime = "F";

        /// <summary>
        /// The format used for all date related items, use this in conjuction with the <see cref="DateFormatCulture"/>
        /// http://msdn.microsoft.com/en-us/library/az4se3k1%28v=vs.110%29.aspx
        /// </summary>
        public const string DataFormat = "MM-dd-yyyy";

        /// <summary>
        /// Normally Outlook will use the subject of a MSG object as it filename. Invalid characters
        /// are replaced by spaces. When there is no subject outlook uses "Nameless" as default
        /// </summary>
        public const string NameLessFileName = "Nameless";

        #region StickyNote constants
        public const string StickyNoteDateLabel = "Date";
        #endregion

        #region Appointment constants
        public const string AppointmentSubjectLabel = "Subject";
        public const string AppointmentLocationLabel = "Location";
        public const string AppointmentStartDateLabel = "Start";
        public const string AppointmentEndDateLabel = "End";
        public const string AppointmentRecurrenceTypeLabel = "Recurrence patern";
        public const string AppointmentRecurrencePaternLabel = "Recurrence patern type";
        public const string AppointmentOrganizerLabel = "Organizer";
        public const string AppointmentMandatoryParticipantsLabel = "Mandatory participants";
        public const string AppointmentOptionalParticipantsLabel = "Optional participants";
        public const string AppointmentCategoriesLabel = "Categories";
        public const string AppointmentResourcesLabel = "Resources";
        public const string AppointmentReccurenceTypeDailyText = "Daily";
        public const string AppointmentReccurenceTypeWeeklyText = "Weekly";
        public const string AppointmentReccurenceTypeMonthlyText = "Monthly";
        public const string AppointmentReccurenceTypeYearlyText = "Yearly";
        public const string AppointmentReccurenceTypeNoneText = "(None)";
        public const string AppointmentClientIntentLabel = "Meeting status";
        public const string AppointmentClientIntentManagerText = "Organizer of the meeting";
        public const string AppointmentClientIntentDelegateText = "Acting as delegate for the organizer of the meeting";
        public const string AppointmentClientIntentDeletedWithNoResponseText = "Deleted meeting with no response";
        public const string AppointmentClientIntentDeletedExceptionWithNoResponseText = "Deleted exception with no response";
        public const string AppointmentClientIntentRespondedTentativeText = "Tentative";
        public const string AppointmentClientIntentRespondedAcceptText = "Accepted";
        public const string AppointmentClientIntentRespondedDeclineText = "Declined";
        public const string AppointmentClientIntentModifiedStartTimeText = "Modified start time";
        public const string AppointmentClientIntentModifiedEndTimeText = "Modified end time";
        public const string AppointmentClientIntentModifiedLocationText = "Modified location";
        public const string AppointmentClientIntentRespondedExceptionDeclineText = "Exception declined";
        public const string AppointmentClientIntentCanceledText = "Declined";
        public const string AppointmentClientIntentExceptionCanceledText = "Declined";
        
        /// <summary>
        /// The appointment ATTACHMENTS label
        /// </summary>
        public const string AppointmentAttachmentsLabel = "Attachments";
        #endregion

        #region Task constants
        public const string TaskSubjectLabel = "Subject";
        public const string TaskStartDateLabel = "Start date";
        public const string TaskDueDateLabel = "Due date";
        public const string TaskDateCompleted = "Completed on";
        public const string TaskStatusLabel = "Status";
        public const string TaskStatusNotStartedText = "Not started";
        public const string TaskStatusInProgressText = "In progress";
        public const string TaskStatusWaitingText = "Waiting";
        public const string TaskStatusCompleteText = "Complete";
        public const string TaskPercentageCompleteLabel = "Percentage complete";
        public const string TaskEstimatedEffortLabel = "Total work";
        public const string TaskActualEffortLabel = "Actual work";
        public const string TaskOwnerLabel = "Owner";
        public const string TaskContactsLabel = "Contacts";
        public const string TaskCompanyLabel = "Company";
        public const string TaskBillingInformationLabel = "Billing information";
        public const string TaskMileageLabel = "Mileage";
        public const string TaskRequestedByLabel = "Requested by";
        #endregion

        #region DateDifference constants
        public const string DateDifferenceYearsText = "years";
        public const string DateDifferenceYearText = "year";

        public const string DateDifferenceMonthsText = "months";
        public const string DateDifferenceMonthText = "month";
        
        public const string DateDifferenceWeeksText = "weeks";
        public const string DateDifferenceWeekText = "week";
        
        public const string DateDifferenceDaysText = "days";
        public const string DateDifferenceDayText = "day";
        
        public const string DateDifferenceHoursText = "hours";
        public const string DateDifferenceHourText = "hour";
        
        public const string DateDifferenceMinutesText = "minutes";
        public const string DateDifferenceMinuteText = "minute";

        public const string DateDifferenceSecondsText = "seconds";
        public const string DateDifferenceSecondText = "second";
        #endregion

        #region Importance
        public const string ImportanceLabel = "Urgent";
        public const string ImportanceLowText = "Low";
        public const string ImportanceNormalText = "Normal";
        public const string ImportanceHighText = "High";
        #endregion

        #region Contact constants
        public const string DisplayNameLabel = "Full name";
        public const string SurNameLabel = "Last name";
        public const string GivenNameLabel = "First name";
        public const string FunctionLabel = "Job title";
        public const string DepartmentLabel = "Department";
        public const string CompanyLabel = "Company";
        public const string WorkAddressLabel = "Business address";
        public const string BusinessTelephoneNumberLabel = "Business";
        public const string BusinessTelephoneNumber2Label = "Business 2";
        public const string BusinessFaxNumberLabel = "Business fax";
        public const string HomeAddressLabel = "Home address";
        public const string HomeTelephoneNumberLabel = "Home";
        public const string HomeTelephoneNumber2Label = "Home 2";
        public const string HomeFaxNumberLabel = "Home fax";
        public const string OtherAddressLabel = "Other address";
        public const string OtherFaxLabel = "Other fax";
        public const string OtherTelephoneNumberLabel = "Other";
        public const string PrimaryTelephoneNumberLabel = "Primary phone";
        public const string PrimaryFaxNumberLabel = "Other fax";
        public const string AssistantLabel = "Assistant";
        public const string AssistantTelephoneNumberLabel = "Assistant";
        public const string InstantMessagingAddressLabel = "IM address";
        public const string CompanyMainTelephoneNumberLabel = "Company main phone";
        public const string CellularTelephoneNumberLabel = "Mobile";
        public const string CarTelephoneNumberLabel = "Car";
        public const string RadioTelephoneNumberLabel = "Radio";
        public const string BeeperTelephoneNumberLabel = "Pager";
        public const string CallbackTelephoneNumberLabel = "Callback";
        public const string TextTelephoneLabel = "TTY/TDD phone";
        public const string ISDNNumberLabel = "ISDN";
        public const string TelexNumberLabel = "Telex";
        public const string Email1EmailAddressLabel = "E-mail";
        public const string Email1DisplayNameLabel = "E-mail display as";
        public const string Email2EmailAddressLabel = "E-mail 2";
        public const string Email2DisplayNameLabel = "E-mail 2 display as";
        public const string Email3EmailAddressLabel = "E-mail 3";
        public const string Email3DisplayNameLabel = "E-mail 3 display as";
        public const string BirthdayLabel = "Birthday";
        public const string WeddingAnniversaryLabel = "Anniversary";
        public const string SpouseNameLabel = "Spouse/Partner";
        public const string ProfessionLabel = "Profession";
        public const string HtmlLabel = "Web page";
        #endregion
    }
}
