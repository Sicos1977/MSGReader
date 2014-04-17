using System;
using System.Collections.ObjectModel;

namespace DocumentServices.Modules.Readers.MsgReader.Outlook
{
    internal partial class Storage
    {        
        /// <summary>
        /// Class used to contain all the task information. A task can also be added to a E-mail (<see cref="Storage.Message"/>) when
        /// the FollowUp flag is set.
        /// </summary>
        public sealed class Task : Storage
        {
            #region Public enum TaskStatus
            public enum TaskStatus
            {
                /// <summary>
                /// The task has not yet started
                /// </summary>
                NotStarted = 0,

                /// <summary>
                /// The task is in progress
                /// </summary>
                InProgess = 1,

                /// <summary>
                /// The task is complete
                /// </summary>
                Complete = 2,

                /// <summary>
                /// The task is waiting on someone else
                /// </summary>
                Waiting = 3
            }
            #endregion

            #region Properties
            /// <summary>
            /// Returns the start datetime of the task, null when not available
            /// </summary>
            public DateTime? StartDate
            {
                get { return GetMapiPropertyDateTime(MapiTags.TaskStartDate); }
            }

            /// <summary>
            /// Returns the due datetime of the task, null when not available
            /// </summary>
            public DateTime? DueDate
            {
                get { return GetMapiPropertyDateTime(MapiTags.TaskDueDate); }    
            }

            /// <summary>
            /// Returns the <see cref="TaskStatus">Status</see> of the task, 
            /// null when not available
            /// </summary>
            public TaskStatus? Status
            {
                get { return (TaskStatus) GetMapiPropertyInt32(MapiTags.TaskStatus); }
            }

            /// <summary>
            /// Returns the <see cref="TaskStatus">Status</see> of the task as a string, 
            /// null when not available
            /// </summary>
            public string StatusText
            {
                get
                {
                    switch (Status)
                    {
                        case TaskStatus.NotStarted:
                            return LanguageConsts.TaskStatusNotStartedText;

                        case TaskStatus.InProgess:
                            return LanguageConsts.TaskStatusInProgressText;

                        case TaskStatus.Waiting:
                            return LanguageConsts.TaskStatusWaitingText;

                        case TaskStatus.Complete:
                            return LanguageConsts.TaskStatusCompleteText;

                    }

                    return null;
                }
            }

            /// <summary>
            /// Returns the estimated effort (in minutes) that is needed for the task, 
            /// null when not available
            /// </summary>
            public double? PercentageComplete
            {
                get { return GetMapiPropertyDouble(MapiTags.PercentComplete); }
            }

            /// <summary>
            /// Returns true when the task has been completed, null when not available
            /// </summary>
            public bool? Complete
            {
                get { return GetMapiPropertyBool(MapiTags.TaskComplete); }    
            }

            /// <summary>
            /// Returns the estimated effort (in minutes) that is needed for the task, 
            /// null when no available
            /// </summary>
            public int? EstimatedEffort
            {
                get { return GetMapiPropertyInt32(MapiTags.TaskEstimatedEffort); }
            }
            
            /// <summary>
            /// Returns the actual effort (in minutes) that is spent on the task, 
            /// null when not available
            /// </summary>
            public int? ActualEffort
            {
                get { return GetMapiPropertyInt32(MapiTags.TaskActualEffort); }
            }

            /// <summary>
            /// Returns the owner of the task, null when not available
            /// </summary>
            public string Owner
            {
                get { return GetMapiPropertyString(MapiTags.Companies); }
            }

            /// <summary>
            /// Returns the contacts of the task, null when not available
            /// </summary>
            public ReadOnlyCollection<string> Contacts 
            {
                get { return GetMapiPropertyStringList(MapiTags.Contacts); }    
            }

            /// <summary>
            /// Returns the name of the company for who the task is done, 
            /// null when not available
            /// </summary>
            public string Companies
            {
                get { return GetMapiPropertyString(MapiTags.Companies); }
            }

            /// <summary>
            /// Returns the billing information for the task, null when not available
            /// </summary>
            public string BillingInformation
            {
                get { return GetMapiPropertyString(MapiTags.Billing); }
            }

            /// <summary>
            /// Returns the mileage that is driven to do the task, null when not available
            /// </summary>
            public string Mileage
            {
                get { return GetMapiPropertyString(MapiTags.TaskComplete); }
            }

            /// <summary>
            /// Returns the datetime when the task was completed, only set when <see cref="Complete"/> is true.
            /// Otherwise null
            /// </summary>
            public DateTime? CompleteTime
            {
                get { return GetMapiPropertyDateTime(MapiTags.PR_FLAG_COMPLETE_TIME); }
            }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Task" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Task(Storage message) : base(message._storage)
            {
                _namedProperties = message._namedProperties;
                GC.SuppressFinalize(message);
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;
            }
            #endregion
        }
    }
}