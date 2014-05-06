namespace DocumentServices.Modules.Readers.MsgReader
{
    /// <summary>
    /// This class contains all the language specific constants
    /// </summary>
    internal static class LanguageConstsNL
    {
         /***************************************************************************************
         *  A note to you, please send me this file when you translate it to your own language. *
         *  You can send the file to : sicos2002@hotmail.com                                    *
         ***************************************************************************************/
        
        /// <summary>
        /// The culture format to use for a date
        /// </summary>
        public const string DateFormatCulture = "nl-NL";

        /// <summary>
        /// The format used for all date/time related items, use this in conjuction with the <see cref="DateFormatCulture"/>
        /// http://msdn.microsoft.com/en-us/library/az4se3k1%28v=vs.110%29.aspx
        /// </summary>
        public const string DataFormatWithTime = "F";

        /// <summary>
        /// The format used for all date related items, use this in conjuction with the <see cref="DateFormatCulture"/>
        /// http://msdn.microsoft.com/en-us/library/az4se3k1%28v=vs.110%29.aspx
        /// </summary>
        public const string DataFormat = "dd-MM-yyyy";

        /// <summary>
        /// Normally Outlook will use the subject of a MSG object as it filename. Invalid characters
        /// are replaced by spaces. When there is no subject outlook uses "Nameless" as default
        /// </summary>
        public const string NameLessFileName = "Naamloos";

        #region E-mail constants
        public const string EmailFromLabel = "Van";
        public const string EmailSentOnLabel = "Verzonden op";
        public const string EmailToLabel = "Aan";
        public const string EmailCcLabel = "CC";
        public const string EmailBccLabel = "BCC";
        public const string EmailSubjectLabel = "Onderwerp";
        public const string EmailAttachmentsLabel = "Bijlagen";
        public const string EmailFollowUpFlag = "Vlag";
        public const string EmailFollowUpLabel = "Opvolginsmarkering";
        public const string EmailFollowUpStatusLabel = "Markeringsstatus";
        public const string EmailFollowUpCompletedText = "Voltooid";
        public const string EmailCategoriesLabel = "Categoriëen";
        #endregion

        #region StickyNote constants
        public const string StickyNoteDateLabel = "Datum";
        #endregion

        #region Appointment constants
        public const string AppointmentSubjectLabel = "Onderwerp";
        public const string AppointmentLocationLabel = "Locatie";
        public const string AppointmentStartDateLabel = "Begin";
        public const string AppointmentEndDateLabel = "Einde";
        public const string AppointmentRecurrenceTypeLabel = "Terugkeerpatroon";
        public const string AppointmentRecurrencePaternLabel = "Type terugkeerpatroon";
        public const string AppointmentOrganizerLabel = "Organisator";
        public const string AppointmentMandatoryParticipantsLabel = "Verplichte deelnemers";
        public const string AppointmentOptionalParticipantsLabel = "Optionele deelnemers";
        public const string AppointmentCategoriesLabel = "Categoriëen";
        public const string AppointmentResourcesLabel = "Resources";
        public const string AppointmentReccurenceTypeDailyText = "Dagelijks";
        public const string AppointmentReccurenceTypeWeeklyText = "Wekelijks";
        public const string AppointmentReccurenceTypeMonthlyText = "Maandelijks";
        public const string AppointmentReccurenceTypeYearlyText = "Jaarlijks";
        public const string AppointmentReccurenceTypeNoneText = "(GEEN)";
        public const string AppointmentClientIntentLabel = "Vergaderingsstatus";
        public const string AppointmentClientIntentManagerText = "organisator van de afspraak";
        public const string AppointmentClientIntentDelegateText = "Acteert als een gedelegeerde van de eigenaar van de afspraak";
        public const string AppointmentClientIntentDeletedWithNoResponseText = "Afspraak geweigerd zonder reactie";
        public const string AppointmentClientIntentDeletedExceptionWithNoResponseText = "Wijziging verwijdert zonder reactie";
        public const string AppointmentClientIntentRespondedTentativeText = "Voorlopig";
        public const string AppointmentClientIntentRespondedAcceptText = "Geaccepteerd";
        public const string AppointmentClientIntentRespondedDeclineText = "Geweigerd";
        public const string AppointmentClientIntentModifiedStartTimeText = "Aangepaste begin tijd";
        public const string AppointmentClientIntentModifiedEndTimeText = "Aangepaste eind tijd";
        public const string AppointmentClientIntentModifiedLocationText = "Aangepaste locatie";
        public const string AppointmentClientIntentRespondedExceptionDeclineText = "Wijziging geweigerd";
        public const string AppointmentClientIntentCanceledText = "Geweigerd";
        public const string AppointmentClientIntentExceptionCanceledText = "Geweigerd";
        
        /// <summary>
        /// The appointment ATTACHMENTS label
        /// </summary>
        public const string AppointmentAttachmentsLabel = "Bijlagen";
        #endregion

        #region Task constants
        public const string TaskSubjectLabel = "Onderwerp";
        public const string TaskStartDateLabel = "Begindatum";
        public const string TaskDueDateLabel = "Einddatum";
        public const string TaskDateCompleted = "Voltooid op";
        public const string TaskStatusLabel = "Status";
        public const string TaskStatusNotStartedText = "Niet gestart";
        public const string TaskStatusInProgressText = "Wordt uitegevoerd";
        public const string TaskStatusWaitingText = "Wacht op iemand anders";
        public const string TaskStatusCompleteText = "Voltooid";
        public const string TaskPercentageCompleteLabel = "Percentage voltooid";
        public const string TaskEstimatedEffortLabel = "Totale hoeveelheid werk";
        public const string TaskActualEffortLabel = "Reële hoeveelheid werk";
        public const string TaskOwnerLabel = "Eigenaar";
        public const string TaskContactsLabel = "Contacten";
        public const string TaskCompanyLabel = "Bedrijf";
        public const string TaskBillingInformationLabel = "Factuur informatie";
        public const string TaskMileageLabel = "Reisafstand";
        public const string TaskRequestedByLabel = "Verzocht door";
        #endregion

        #region DateDifference constants
        public const string DateDifferenceYearsText = "jaren";
        public const string DateDifferenceYearText = "jaar";

        public const string DateDifferenceMonthsText = "maanden";
        public const string DateDifferenceMonthText = "maand";
        
        public const string DateDifferenceWeeksText = "weken";
        public const string DateDifferenceWeekText = "week";
        
        public const string DateDifferenceDaysText = "dagen";
        public const string DateDifferenceDayText = "dag";
        
        public const string DateDifferenceHoursText = "uren";
        public const string DateDifferenceHourText = "uur";
        
        public const string DateDifferenceMinutesText = "minuten";
        public const string DateDifferenceMinuteText = "minuut";

        public const string DateDifferenceSecondsText = "seconden";
        public const string DateDifferenceSecondText = "seconde";
        #endregion

        #region Importance
        public const string ImportanceLabel = "Prioriteit";
        public const string ImportanceLowText = "Laag";
        public const string ImportanceNormalText = "";
        public const string ImportanceHighText = "Hoog";
        #endregion

        #region Contact constants
        public const string DisplayNameLabel = "Volledige naam";
        public const string SurNameLabel = "Achternaam";
        public const string GivenNameLabel = "Voornaam";
        public const string FunctionLabel = "Functie";
        public const string DepartmentLabel = "Afdeling";
        public const string CompanyLabel = "Bedrijf";
        public const string WorkAddressLabel = "Werkadres";
        public const string BusinessTelephoneNumberLabel = "Werk";
        public const string BusinessTelephoneNumber2Label = "Werk 2";
        public const string BusinessFaxNumberLabel = "Fax op werk";
        public const string HomeAddressLabel = "Huisadres";
        public const string HomeTelephoneNumberLabel = "Thuis";
        public const string HomeTelephoneNumber2Label = "Thuis 2";
        public const string HomeFaxNumberLabel = "Fax thuis";
        public const string OtherAddressLabel = "Ander adres";
        public const string OtherFaxLabel = "Andere fax";
        public const string OtherTelephoneNumberLabel = "Overige";
        public const string PrimaryTelephoneNumberLabel = "Hoofdtelefoon";
        public const string PrimaryFaxNumberLabel = "Andere fax";
        public const string AssistantLabel = "Assistent";
        public const string AssistantTelephoneNumberLabel = "Assistent";
        public const string InstantMessagingAddressLabel = "IM-adres";
        public const string CompanyMainTelephoneNumberLabel = "Hoofdtelefoon bedrijf";
        public const string CellularTelephoneNumberLabel = "Mobiel";
        public const string CarTelephoneNumberLabel = "Auto";
        public const string RadioTelephoneNumberLabel = "Radio";
        public const string BeeperTelephoneNumberLabel = "Pager";
        public const string CallbackTelephoneNumberLabel = "Terugbellen";
        public const string TextTelephoneLabel = "Teksttelefoon";
        public const string ISDNNumberLabel = "ISDN";
        public const string TelexNumberLabel = "Telex";
        public const string Email1EmailAddressLabel = "E-mail";
        public const string Email1DisplayNameLabel = "E-mail - weergeven naam1";
        public const string Email2EmailAddressLabel = "E-mail 2";
        public const string Email2DisplayNameLabel = "E-mail - weergeven naam2";
        public const string Email3EmailAddressLabel = "E-mail 3";
        public const string Email3DisplayNameLabel = "E-mail - weergeven naam3";
        public const string BirthdayLabel = "Verjaardag";
        public const string WeddingAnniversaryLabel = "Speciale datum";
        public const string SpouseNameLabel = "Partner";
        public const string ProfessionLabel = "Beroek";
        public const string HtmlLabel = "Web pagina";
        #endregion
    }
}
