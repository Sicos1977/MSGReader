using System;

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
            /// Returns the start datetime of the task
            /// </summary>
            public DateTime? StartDate
            {
                get { return GetMapiPropertyDateTime(MapiTags.TaskStartDate); }
            }

            /// <summary>
            /// Returns the due datetime of the task
            /// </summary>
            public DateTime? DueDate
            {
                get { return GetMapiPropertyDateTime(MapiTags.TaskDueDate); }    
            }

            /// <summary>
            /// Returns the <see cref="TaskStatus">Status</see> of the task
            /// </summary>
            public TaskStatus? Status
            {
                get { return (TaskStatus) GetMapiPropertyInt32(MapiTags.TaskStatus); }
            }

            /// <summary>
            /// Returns true when the task has been completed
            /// </summary>
            public bool? Complete
            {
                get { return GetMapiPropertyBool(MapiTags.TaskComplete); }    
            }

            /// <summary>
            /// Returns the datetime when the task was completed, only set when <see cref="Complete"/> is true
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