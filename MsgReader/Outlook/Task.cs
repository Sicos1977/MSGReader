using System;
using System.Collections.ObjectModel;
using MsgReader.Helpers;
using MsgReader.Localization;

/*
   Copyright 2013-2017 Kees van Spelde

   Licensed under The Code Project Open License (CPOL) 1.02;
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.codeproject.com/info/cpol10.aspx

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

namespace MsgReader.Outlook
{
    public partial class Storage
    {        
        /// <summary>
        /// Class used to contain all the task information. A task can also be added to an E-mail (<see cref="Storage.Message"/>) when
        /// the FollowUp flag is set.
        /// </summary>
        public sealed class Task : Storage
        {
            #region Public enum TaskStatus
            /// <summary>
            /// The status of a task
            /// </summary>
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
            /// Returns the start datetime of the <see cref="Storage.Task"/>, null when not available
            /// </summary>
            public DateTime? StartDate { get; private set ; }

            /// <summary>
            /// Returns the due datetime of the <see cref="Storage.Task"/>, null when not available
            /// </summary>
            public DateTime? DueDate { get; private set; }

            /// <summary>
            /// Returns the <see cref="TaskStatus">Status</see> of the <see cref="Storage.Task"/>, 
            /// null when not available
            /// </summary>
            public TaskStatus? Status { get; private set; }

            /// <summary>
            /// Returns the <see cref="TaskStatus">Status</see> of the <see cref="Storage.Task"/> as a string, 
            /// null when not available
            /// </summary>
            public string StatusText { get; private set; }

            /// <summary>
            /// Returns the estimated effort (in minutes) that is needed for <see cref="Storage.Task"/> task, 
            /// null when not available
            /// </summary>
            public double? PercentageComplete { get; private set; }

            /// <summary>
            /// Returns true when the <see cref="Storage.Task"/> has been completed, null when not available
            /// </summary>
            public bool? Complete { get; private set; }

            /// <summary>
            /// Returns the estimated effort that is needed for the <see cref="Storage.Task"/> as a <see cref="TimeSpan"/>, 
            /// null when no available
            /// </summary>
            public TimeSpan? EstimatedEffort { get; private set; }

            /// <summary>
            /// Returns the estimated effort that is needed for the <see cref="Storage.Task"/> as a string (e.g. 11 weeks), 
            /// null when no available
            /// </summary>
            public string EstimatedEffortText { get; private set; }
            
            /// <summary>
            /// Returns the actual effort that is spent on the <see cref="Storage.Task"/> as a <see cref="TimeSpan"/>,
            /// null when not available
            /// </summary>
            public TimeSpan? ActualEffort { get; private set; }

            /// <summary>
            /// Returns the actual effort that is spent on the <see cref="Storage.Task"/> as a string (e.g. 11 weeks), 
            /// null when no available
            /// </summary>
            public string ActualEffortText { get; private set; }

            /// <summary>
            /// Returns the owner of the <see cref="Storage.Task"/>, null when not available
            /// </summary>
            public string Owner { get; private set; }

            /// <summary>
            /// Returns the contacts of the <see cref="Storage.Task"/>, null when not available
            /// </summary>
            public ReadOnlyCollection<string> Contacts { get; private set; }

            /// <summary>
            /// Returns the name of the company for who the task is done, 
            /// null when not available
            /// </summary>
            public ReadOnlyCollection<string> Companies { get; private set; }

            /// <summary>
            /// Returns the billing information for the <see cref="Storage.Task"/>, null when not available
            /// </summary>
            public string BillingInformation { get; private set; }

            /// <summary>
            /// Returns the mileage that is driven to do the <see cref="Storage.Task"/>, null when not available
            /// </summary>
            public string Mileage { get; private set; }

            /// <summary>
            /// Returns the datetime when the <see cref="Storage.Task"/> was completed, 
            /// only set when <see cref="Complete"/> is true.
            /// Otherwise null
            /// </summary>
            public DateTime? CompleteTime { get; private set; }
            #endregion

            #region Constructor
            /// <summary>
            /// Initializes a new instance of the <see cref="Storage.Task" /> class.
            /// </summary>
            /// <param name="message"> The message. </param>
            internal Task(Storage message) : base(message._storage)
            {
                //GC.SuppressFinalize(message);
                _namedProperties = message._namedProperties;
                _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

                StartDate = GetMapiPropertyDateTime(MapiTags.TaskStartDate);
                DueDate = GetMapiPropertyDateTime(MapiTags.TaskDueDate);

                var status = GetMapiPropertyInt32(MapiTags.TaskStatus);
                if (status == null)
                    Status = null;
                else
                    Status = (TaskStatus)status;

                switch (Status)
                {
                    case TaskStatus.NotStarted:
                        StatusText = LanguageConsts.TaskStatusNotStartedText;
                        break;

                    case TaskStatus.InProgess:
                        StatusText = LanguageConsts.TaskStatusInProgressText;
                        break;

                    case TaskStatus.Waiting:
                        StatusText = LanguageConsts.TaskStatusWaitingText;
                        break;

                    case TaskStatus.Complete:
                        StatusText = LanguageConsts.TaskStatusCompleteText;
                        break;

                    default:
                        StatusText = null;
                        break;
                }

                PercentageComplete = GetMapiPropertyDouble(MapiTags.PercentComplete);
                Complete = GetMapiPropertyBool(MapiTags.TaskComplete);

                var estimatedEffort = GetMapiPropertyInt32(MapiTags.TaskEstimatedEffort);
                if (estimatedEffort == null)
                    EstimatedEffort = null;
                else
                    EstimatedEffort = new TimeSpan(0, 0, (int)estimatedEffort);

                var now = DateTime.Now;

                EstimatedEffortText = (EstimatedEffort == null
                    ? null
                    : DateDifference.Difference(now, now + ((TimeSpan) EstimatedEffort)).ToString());

                var actualEffort = GetMapiPropertyInt32(MapiTags.TaskActualEffort);
                if (actualEffort == null)
                    ActualEffort = null; 
                else 
                    ActualEffort = new TimeSpan(0, 0, (int) actualEffort);

                ActualEffortText = (ActualEffort == null
                    ? null
                    : DateDifference.Difference(now, now + ((TimeSpan) ActualEffort)).ToString());

                Owner = GetMapiPropertyString(MapiTags.Owner);
                Contacts = GetMapiPropertyStringList(MapiTags.Contacts);
                Companies = GetMapiPropertyStringList(MapiTags.Companies); 
                BillingInformation = GetMapiPropertyString(MapiTags.Billing);
                Mileage = GetMapiPropertyString(MapiTags.Mileage);
                CompleteTime = GetMapiPropertyDateTime(MapiTags.PR_FLAG_COMPLETE_TIME);
            }
            #endregion
        }
    }
}