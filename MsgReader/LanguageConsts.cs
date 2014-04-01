namespace DocumentServices.Modules.Readers.MsgReader
{
    /// <summary>
    /// This class contains all the language specific constants
    /// </summary>
    internal static class LanguageConsts
    {
        /// <summary>
        /// The sender of the MSG
        /// </summary>
        public const string FromLabel = "From";

        /// <summary>
        /// When the MSG object has been sent by the originator
        /// </summary>
        public const string SentOnLabel = "Sent on";

        /// <summary>
        /// The recipient of the MSG object
        /// </summary>
        public const string ToLabel = "To";

        /// <summary>
        /// The CC's of the MSG object
        /// </summary>
        public const string CcLabel = "CC";

        /// <summary>
        /// The BCC's of the MSG object
        /// </summary>
        public const string BccLabel = "BCC";

        //public const string receivedOnLabel = "Received on";

        /// <summary>
        /// The subject of the MSG object
        /// </summary>
        public const string SubjectLabel = "Subject";

        /// <summary>
        /// The label for the attachments in the MSG object
        /// </summary>
        public const string AttachmentsLabel = "Attachments";

        /// <summary>
        /// The follow up flag
        /// </summary>
        public const string FollowUpFlag = "Flag";

        /// <summary>
        /// The text for the follow up flag
        /// </summary>
        public const string FollowUpLabel = "Follow up";

        /// <summary>
        /// The text for the follow up status
        /// </summary>
        public const string FollowUpStatusLabel = "Follow up status";

        /// <summary>
        /// Text to show when the follow up flag has been set to completed
        /// </summary>
        public const string FollowUpCompletedText = "Completed";

        /// <summary>
        /// The task start date label
        /// </summary>
        public const string TaskStartDateLabel = "Startdate";

        /// <summary>
        /// The task duedate label
        /// </summary>
        public const string TaskDueDateLabel = "Enddate";

        /// <summary>
        /// The date when the task has been set to completed
        /// </summary>
        public const string TaskDateCompleted = "Completed on";

        /// <summary>
        /// The categories label
        /// </summary>
        public const string CategoriesLabel = "Categories";

        /// <summary>
        /// The format used for all date related items
        /// </summary>
        public const string DataFormat = "dd-MM-yyyy HH:mm:ss";

        /// <summary>
        /// Normally Outlook will use the subject of a MSG object as it filename. Invalid characters
        /// are replaced by spaces. When there is no subject outlook uses "Nameless" as default
        /// </summary>
        public const string NameLessFileName = "Nameless";

        /// <summary>
        /// The sticky note date label 
        /// </summary>
        public const string StickyNoteDate = "Date";
    }
}
