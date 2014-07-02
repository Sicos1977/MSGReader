namespace DocumentServices.Modules.Readers.MsgReader
{
    /// <summary>
    /// This class contains all the language specific constants
    /// </summary>
    internal static class LanguageConstsDE
    {
         /***************************************************************************************
         * Translated to German by Ronald Kohl                                                  *
         ***************************************************************************************/

        /// <summary>
        /// The culture format to use for a date
        /// </summary>
        public const string DateFormatCulture = "de-DE";

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
        /// Normally Outlook will use the subject of a MSG object as it's filename. Invalid characters
        /// are replaced by spaces. When there is no subject outlook uses "Nameless" as default
        /// </summary>
        public const string NameLessFileName = "Namenlos";
        
        #region E-mail constants
        public const string EmailFromLabel = "Von";
        public const string EmailSentOnLabel = "Gesendet am";
        public const string EmailToLabel = "An";
        public const string EmailCcLabel = "CC";
        public const string EmailBccLabel = "BCC";
        public const string EmailSubjectLabel = "Betreff";
        public const string EmailAttachmentsLabel = "Attachments";
        public const string EmailFollowUpFlag = "Flag";
        public const string EmailFollowUpLabel = "Verlaufsmarkierung";
        public const string EmailFollowUpStatusLabel = "Verlaufsstatus";
        public const string EmailFollowUpCompletedText = "Verlauf komplett";
        public const string EmailCategoriesLabel = "Kategorien";
        #endregion

        #region StickyNote constants
        public const string StickyNoteDateLabel = "Datum";
        #endregion

        #region Appointment constants
        public const string AppointmentSubjectLabel = "Betreff";
        public const string AppointmentLocationLabel = "Ort";
        public const string AppointmentStartDateLabel = "Beginn";
        public const string AppointmentEndDateLabel = "Ende";
        public const string AppointmentRecurrenceTypeLabel = "Serientyp";
        public const string AppointmentRecurrencePaternLabel = "Recurrence patern type";
        public const string AppointmentOrganizerLabel = "Organisator";
        public const string AppointmentMandatoryParticipantsLabel = "Teilnehmer";
        public const string AppointmentOptionalParticipantsLabel = "Optionale Teilnehmer";
        public const string AppointmentCategoriesLabel = "Kategorien";
        public const string AppointmentResourcesLabel = "Ressourcen";
        public const string AppointmentReccurenceTypeDailyText = "Täglich";
        public const string AppointmentReccurenceTypeWeeklyText = "Wöchentlich";
        public const string AppointmentReccurenceTypeMonthlyText = "Monatlich";
        public const string AppointmentReccurenceTypeYearlyText = "Jährlich";
        public const string AppointmentReccurenceTypeNoneText = "(Keine)";
        public const string AppointmentClientIntentLabel = "Status des Meeting";
        public const string AppointmentClientIntentManagerText = "Organisator des Meetings";
        public const string AppointmentClientIntentDelegateText = "Vertreter des Organisators des Meetings";
        public const string AppointmentClientIntentDeletedWithNoResponseText = "Meeting ohne Antwort gelöscht";
        public const string AppointmentClientIntentDeletedExceptionWithNoResponseText = "Gelöschte Ausnahme ohne Antwort";
        public const string AppointmentClientIntentRespondedTentativeText = "Vorläufig";
        public const string AppointmentClientIntentRespondedAcceptText = "Akzeptiert";
        public const string AppointmentClientIntentRespondedDeclineText = "Abgelehnt";
        public const string AppointmentClientIntentModifiedStartTimeText = "Geänderter Beginn";
        public const string AppointmentClientIntentModifiedEndTimeText = "Geändertes Ende";
        public const string AppointmentClientIntentModifiedLocationText = "Geänderter Ort";
        public const string AppointmentClientIntentRespondedExceptionDeclineText = "Ausnahme abgelehnt";
        public const string AppointmentClientIntentCanceledText = "Abgelehnt";
        public const string AppointmentClientIntentExceptionCanceledText = "Abgelehnt";
        public const string AppointmentAttachmentsLabel = "Attachments";
        #endregion

        #region Task constants
        public const string TaskSubjectLabel = "Betreff";
        public const string TaskStartDateLabel = "Beginn";
        public const string TaskDueDateLabel = "Dauer";
        public const string TaskDateCompleted = "Fällig am";
        public const string TaskStatusLabel = "Status";
        public const string TaskStatusNotStartedText = "Nicht begonnen";
        public const string TaskStatusInProgressText = "In Bearbeitung";
        public const string TaskStatusWaitingText = "Wartet";
        public const string TaskStatusCompleteText = "Erledigt";
        public const string TaskPercentageCompleteLabel = "Erledigt in Prozent";
        public const string TaskEstimatedEffortLabel = "Gesamtaufwand";
        public const string TaskActualEffortLabel = "IST-Arbeit";
        public const string TaskOwnerLabel = "Eigentümer";
        public const string TaskContactsLabel = "Kontakt";
        public const string TaskCompanyLabel = "Firma";
        public const string TaskBillingInformationLabel = "Abrechungsinformation";
        public const string TaskMileageLabel = "Reisekilometer";
        public const string TaskRequestedByLabel = "Angefordert von";
        #endregion

        #region DateDifference constants
        public const string DateDifferenceYearsText = "Jahre";
        public const string DateDifferenceYearText = "Jahr";

        public const string DateDifferenceMonthsText = "Monate";
        public const string DateDifferenceMonthText = "Monat";
        
        public const string DateDifferenceWeeksText = "Wochen";
        public const string DateDifferenceWeekText = "Woche";
        
        public const string DateDifferenceDaysText = "Tage";
        public const string DateDifferenceDayText = "Tag";
        
        public const string DateDifferenceHoursText = "Stunden";
        public const string DateDifferenceHourText = "Stunde";
        
        public const string DateDifferenceMinutesText = "Minuten";
        public const string DateDifferenceMinuteText = "Minute";

        public const string DateDifferenceSecondsText = "Sekunden";
        public const string DateDifferenceSecondText = "Sekunde";
        #endregion

        #region Importance
        public const string ImportanceLabel = "Priorität";
        public const string ImportanceLowText = "Niedrig";
        public const string ImportanceNormalText = "";
        public const string ImportanceHighText = "Hoch";
        #endregion

        #region Contact constants
        public const string DisplayNameLabel = "Full name";
        public const string SurNameLabel = "Nachname";
        public const string GivenNameLabel = "Vorname";
        public const string FunctionLabel = "Job title";
        public const string DepartmentLabel = "Abteilung";
        public const string CompanyLabel = "Firma";
        public const string WorkAddressLabel = "Geschäftsadresse";
        public const string BusinessTelephoneNumberLabel = "Telefon Firma";
        public const string BusinessTelephoneNumber2Label = "Telefon 2 Firma";
        public const string BusinessFaxNumberLabel = "Fax Firma";
        public const string HomeAddressLabel = "Hausanschrift";
        public const string HomeTelephoneNumberLabel = "Telefon privat";
        public const string HomeTelephoneNumber2Label = "Telefon 2 privat";
        public const string HomeFaxNumberLabel = "Fax privat";
        public const string OtherAddressLabel = "Weitere Adresse";
        public const string OtherFaxLabel = "Weiteres Fax";
        public const string OtherTelephoneNumberLabel = "Weiters Telefon";
        public const string PrimaryTelephoneNumberLabel = "Primäres Telefon";
        public const string PrimaryFaxNumberLabel = "Primäres Fax";
        public const string AssistantLabel = "Assistenz";
        public const string AssistantTelephoneNumberLabel = "Assistenz";
        public const string InstantMessagingAddressLabel = "Instant Messaging Adresse";
        public const string CompanyMainTelephoneNumberLabel = "Firma Haupttelefon";
        public const string CellularTelephoneNumberLabel = "Mobile";
        public const string CarTelephoneNumberLabel = "Autotelefon";
        public const string RadioTelephoneNumberLabel = "Funk";
        public const string BeeperTelephoneNumberLabel = "Pager";
        public const string CallbackTelephoneNumberLabel = "Rückrufnummer";
        public const string TextTelephoneLabel = "TTY/TDD phone";
        public const string ISDNNumberLabel = "ISDN";
        public const string TelexNumberLabel = "Telex";
        public const string Email1EmailAddressLabel = "E-mail";
        public const string Email1DisplayNameLabel = "E-mail Anzeige als";
        public const string Email2EmailAddressLabel = "E-mail 2";
        public const string Email2DisplayNameLabel = "E-mail 2 Anzeige als";
        public const string Email3EmailAddressLabel = "E-mail 3";
        public const string Email3DisplayNameLabel = "E-mail 3 Anzeige als";
        public const string BirthdayLabel = "Geburtstag";
        public const string WeddingAnniversaryLabel = "Jahrestag";
        public const string SpouseNameLabel = "Ehepartner/Partner";
        public const string ProfessionLabel = "Beruf";
        public const string HtmlLabel = "Web page";
        #endregion
    }
}
