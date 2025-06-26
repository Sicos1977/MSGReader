﻿//
// Task.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2024 Kees van Spelde. (www.magic-sessions.com)
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
using System.Collections.ObjectModel;
using MsgReader.Helpers;
using MsgReader.Localization;

namespace MsgReader.Outlook;

#region Enum TaskStatus
/// <summary>
///     The status of a task
/// </summary>
public enum TaskStatus
{
    /// <summary>
    ///     The task has not yet started
    /// </summary>
    NotStarted = 0,

    /// <summary>
    ///     The task is in progress
    /// </summary>
    InProgress = 1,

    /// <summary>
    ///     The task is complete
    /// </summary>
    Complete = 2,

    /// <summary>
    ///     The task is waiting on someone else
    /// </summary>
    Waiting = 3
}
#endregion

public partial class Storage
{
    /// <summary>
    ///     Class used to contain all the task information. A task can also be added to an E-mail (
    ///     <see cref="Storage.Message" />) when
    ///     the FollowUp flag is set.
    /// </summary>
    public sealed class Task : Storage
    {
        #region Properties
        /// <summary>
        ///     Returns the start datetime of the <see cref="Storage.Task" />, null when not available
        /// </summary>
        public DateTimeOffset? StartDate { get; }

        /// <summary>
        ///     Returns the due datetime of the <see cref="Storage.Task" />, null when not available
        /// </summary>
        public DateTimeOffset? DueDate { get; }

        /// <summary>
        ///     Returns the <see cref="TaskStatus">Status</see> of the <see cref="Storage.Task" />,
        ///     null when not available
        /// </summary>
        public TaskStatus? Status { get; }

        /// <summary>
        ///     Returns the <see cref="TaskStatus">Status</see> of the <see cref="Storage.Task" /> as a string,
        ///     null when not available
        /// </summary>
        public string StatusText { get; }

        /// <summary>
        ///     Returns the estimated effort (in minutes) that is needed for <see cref="Storage.Task" /> task,
        ///     null when not available
        /// </summary>
        public double? PercentageComplete { get; }

        /// <summary>
        ///     Returns true when the <see cref="Storage.Task" /> has been completed, null when not available
        /// </summary>
        public bool? Complete { get; }

        /// <summary>
        ///     Returns the estimated effort that is needed for the <see cref="Storage.Task" /> as a <see cref="TimeSpan" />,
        ///     null when no available
        /// </summary>
        public TimeSpan? EstimatedEffort { get; }

        /// <summary>
        ///     Returns the estimated effort that is needed for the <see cref="Storage.Task" /> as a string (e.g. 11 weeks),
        ///     null when no available
        /// </summary>
        public string EstimatedEffortText { get; }

        /// <summary>
        ///     Returns the actual effort that is spent on the <see cref="Storage.Task" /> as a <see cref="TimeSpan" />,
        ///     null when not available
        /// </summary>
        public TimeSpan? ActualEffort { get; }

        /// <summary>
        ///     Returns the actual effort that is spent on the <see cref="Storage.Task" /> as a string (e.g. 11 weeks),
        ///     null when no available
        /// </summary>
        public string ActualEffortText { get; }

        /// <summary>
        ///     Returns the owner of the <see cref="Storage.Task" />, null when not available
        /// </summary>
        public string Owner { get; }

        /// <summary>
        ///     Returns the contacts of the <see cref="Storage.Task" />, null when not available
        /// </summary>
        public ReadOnlyCollection<string> Contacts { get; }

        /// <summary>
        ///     Returns the name of the company for whom the task is done,
        ///     null when not available
        /// </summary>
        public ReadOnlyCollection<string> Companies { get; }

        /// <summary>
        ///     Returns the billing information for the <see cref="Storage.Task" />, null when not available
        /// </summary>
        public string BillingInformation { get; }

        /// <summary>
        ///     Returns the mileage that is driven to do the <see cref="Storage.Task" />, null when not available
        /// </summary>
        public string Mileage { get; }

        /// <summary>
        ///     Returns the <see cref="DateTimeOffset"/>> when the <see cref="Storage.Task" /> was completed,<br/>
        ///     only set when <see cref="Complete" /> is <c>true</c>, otherwise null
        /// </summary>
        /// <remarks>
        ///     Use <see cref="DateTimeOffset.ToLocalTime"/> to get the local time
        /// </remarks>
        public DateTimeOffset? CompleteTime { get; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Initializes a new instance of the <see cref="Storage.Task" /> class.
        /// </summary>
        /// <param name="message"> The message. </param>
        internal Task(Storage message) : base(message._rootStorage)
        {
            _namedProperties = message._namedProperties;
            _propHeaderSize = MapiTags.PropertiesStreamHeaderTop;

            StartDate = GetMapiPropertyDateTimeOffset(MapiTags.TaskStartDate);
            DueDate = GetMapiPropertyDateTimeOffset(MapiTags.TaskDueDate);

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

                case TaskStatus.InProgress:
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

            EstimatedEffortText = EstimatedEffort == null
                ? null
                : DateDifference.Difference(now, now + (TimeSpan)EstimatedEffort).ToString();

            var actualEffort = GetMapiPropertyInt32(MapiTags.TaskActualEffort);
            if (actualEffort == null)
                ActualEffort = null;
            else
                ActualEffort = new TimeSpan(0, 0, (int)actualEffort);

            ActualEffortText = ActualEffort == null
                ? null
                : DateDifference.Difference(now, now + (TimeSpan)ActualEffort).ToString();

            Owner = GetMapiPropertyString(MapiTags.Owner);
            Contacts = GetMapiPropertyStringList(MapiTags.Contacts);
            Companies = GetMapiPropertyStringList(MapiTags.Companies);
            BillingInformation = GetMapiPropertyString(MapiTags.Billing);
            Mileage = GetMapiPropertyString(MapiTags.Mileage);
            CompleteTime = GetMapiPropertyDateTimeOffset(MapiTags.PR_FLAG_COMPLETE_TIME);
        }
        #endregion
    }
}