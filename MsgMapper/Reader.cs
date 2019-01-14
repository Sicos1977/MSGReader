//
// Reader.cs
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
using System.IO;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using MsgReader.Outlook;

namespace MsgMapper
{
    /// <summary>
    /// This class can be used to read an Outlook msg file and map all the known msg properties to
    /// the Windows extended file properties. Windows 7 or up is needed for this
    /// </summary>
    internal class Reader
    {
        #region SetExtendedFileAttributesWithMsgProperties
        /// <summary>
        /// This function will read all the properties of an <see cref="Storage.Message"/> file and maps
        /// all the properties that are filled to the extended file attributes. 
        /// </summary>
        /// <param name="inputFile">The msg file</param>
        public void SetExtendedFileAttributesWithMsgProperties(string inputFile)
        {
            MemoryStream memoryStream = null;

            try
            {
                // We need to read the msg file into memory because we otherwise can't set the extended filesystem
                // properties because the files is locked
                memoryStream = new MemoryStream();
                using (var fileStream = File.OpenRead(inputFile))
                    fileStream.CopyTo(memoryStream);

                memoryStream.Position = 0;

                using (var shellFile = ShellFile.FromFilePath(inputFile))
                {
                    using (var propertyWriter = shellFile.Properties.GetPropertyWriter())
                    {
                        using (var message = new Storage.Message(memoryStream))
                        {
                            switch (message.Type)
                            {
                                case MessageType.Email:
                                    MapEmailPropertiesToExtendedFileAttributes(message, propertyWriter);
                                    break;

                                case MessageType.AppointmentRequest:
                                case MessageType.Appointment:
                                case MessageType.AppointmentResponse:
                                    MapAppointmentPropertiesToExtendedFileAttributes(message, propertyWriter);
                                    break;

                                case MessageType.Task:
                                case MessageType.TaskRequestAccept:
                                    MapTaskPropertiesToExtendedFileAttributes(message, propertyWriter);
                                    break;

                                case MessageType.Contact:
                                    MapContactPropertiesToExtendedFileAttributes(message, propertyWriter);
                                    break;

                                case MessageType.Unknown:
                                    throw new NotSupportedException("Unsupported message type");
                            }
                        }
                    }
                }
            }
            finally
            {
                if (memoryStream != null)
                    memoryStream.Dispose();
            }
        }
        #endregion

        #region MapEmailPropertiesToExtendedFileAttributes
        /// <summary>
        /// Maps all the filled <see cref="Storage.Message"/> properties to the corresponding extended file attributes
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="propertyWriter">The <see cref="ShellPropertyWriter"/> object</param>
        private void MapEmailPropertiesToExtendedFileAttributes(Storage.Message message, ShellPropertyWriter propertyWriter)
        {
            // From
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromName, message.Sender.DisplayName);

            // Sent on
            propertyWriter.WriteProperty(SystemProperties.System.Message.DateSent, message.SentOn);

            // To
            propertyWriter.WriteProperty(SystemProperties.System.Message.ToAddress,
                message.GetEmailRecipients(RecipientType.To, false, false));

            // CC
            propertyWriter.WriteProperty(SystemProperties.System.Message.CcAddress,
                message.GetEmailRecipients(RecipientType.Cc, false, false));

            // BCC
            propertyWriter.WriteProperty(SystemProperties.System.Message.BccAddress,
                message.GetEmailRecipients(RecipientType.Bcc, false, false));

            // Subject
            propertyWriter.WriteProperty(SystemProperties.System.Subject, message.Subject);

            // Urgent
            propertyWriter.WriteProperty(SystemProperties.System.Importance, message.Importance);
            propertyWriter.WriteProperty(SystemProperties.System.ImportanceText, message.ImportanceText);

            // Attachments
            var attachments = message.GetAttachmentNames();
            if (string.IsNullOrEmpty(attachments))
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, false);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, null);
            }
            else
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, true);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, attachments);
            }

            // Clear properties
            propertyWriter.WriteProperty(SystemProperties.System.StartDate, null);
            propertyWriter.WriteProperty(SystemProperties.System.DueDate, null);
            propertyWriter.WriteProperty(SystemProperties.System.DateCompleted, null);
            propertyWriter.WriteProperty(SystemProperties.System.IsFlaggedComplete, null);
            propertyWriter.WriteProperty(SystemProperties.System.FlagStatusText, null);

            // Follow up
            if (message.Flag != null)
            {
                propertyWriter.WriteProperty(SystemProperties.System.IsFlagged, true);
                propertyWriter.WriteProperty(SystemProperties.System.FlagStatusText, message.Flag.Request);

                // Flag status text
                propertyWriter.WriteProperty(SystemProperties.System.FlagStatusText, message.Task.StatusText);

                // When complete
                if (message.Task.Complete != null && (bool)message.Task.Complete)
                {
                    // Flagged complete
                    propertyWriter.WriteProperty(SystemProperties.System.IsFlaggedComplete, true);

                    // Task completed date
                    if (message.Task.CompleteTime != null)
                        propertyWriter.WriteProperty(SystemProperties.System.DateCompleted, (DateTime)message.Task.CompleteTime);
                }
                else
                {
                    // Flagged not complete
                    propertyWriter.WriteProperty(SystemProperties.System.IsFlaggedComplete, false);

                    propertyWriter.WriteProperty(SystemProperties.System.DateCompleted, null);

                    // Task startdate
                    if (message.Task.StartDate != null)
                        propertyWriter.WriteProperty(SystemProperties.System.StartDate, (DateTime)message.Task.StartDate);

                    // Task duedate
                    if (message.Task.DueDate != null)
                        propertyWriter.WriteProperty(SystemProperties.System.DueDate, (DateTime)message.Task.DueDate);
                }
            }

            // Categories
            var categories = message.Categories;
            if (categories != null)
                propertyWriter.WriteProperty(SystemProperties.System.Category, String.Join("; ", String.Join("; ", categories)));
        }
        #endregion

        #region MapAppointmentPropertiesToExtendedFileAttributes
        /// <summary>
        /// Maps all the filled <see cref="Storage.Appointment"/> properties to the corresponding extended file attributes
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="propertyWriter">The <see cref="ShellPropertyWriter"/> object</param>
        private void MapAppointmentPropertiesToExtendedFileAttributes(Storage.Message message, ShellPropertyWriter propertyWriter)
        {
            // From
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromName, message.Sender.DisplayName);

            // Sent on
            if (message.SentOn != null)
                propertyWriter.WriteProperty(SystemProperties.System.Message.DateSent, message.SentOn);

            // Subject
            propertyWriter.WriteProperty(SystemProperties.System.Subject, message.Subject);

            // Location
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.Location, message.Appointment.Location);

            // Start
            propertyWriter.WriteProperty(SystemProperties.System.StartDate, message.Appointment.Start);

            // End
            propertyWriter.WriteProperty(SystemProperties.System.StartDate, message.Appointment.End);

            // Recurrence type
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.IsRecurring,
                message.Appointment.ReccurrenceType != AppointmentRecurrenceType.None);

            // Status
            propertyWriter.WriteProperty(SystemProperties.System.Status, message.Appointment.ClientIntentText);

            // Appointment organizer (FROM)
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.OrganizerAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.OrganizerName, message.Sender.DisplayName);

            // Mandatory participants (TO)
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.RequiredAttendeeNames, message.Appointment.ToAttendees);

            // Optional participants (CC)
            propertyWriter.WriteProperty(SystemProperties.System.Calendar.OptionalAttendeeNames, message.Appointment.CcAttendees);

            // Categories
            var categories = message.Categories;
            if (categories != null)
                propertyWriter.WriteProperty(SystemProperties.System.Category, String.Join("; ", String.Join("; ", categories)));

            // Urgent
            propertyWriter.WriteProperty(SystemProperties.System.Importance, message.Importance);
            propertyWriter.WriteProperty(SystemProperties.System.ImportanceText, message.ImportanceText);

            // Attachments
            var attachments = message.GetAttachmentNames();
            if (string.IsNullOrEmpty(attachments))
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, false);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, null);
            }
            else
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, true);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, attachments);
            }
        }
        #endregion

        #region MapTaskPropertiesToExtendedFileAttributes
        /// <summary>
        /// Maps all the filled <see cref="Storage.Task"/> properties to the corresponding extended file attributes
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="propertyWriter">The <see cref="ShellPropertyWriter"/> object</param>
        private void MapTaskPropertiesToExtendedFileAttributes(Storage.Message message, ShellPropertyWriter propertyWriter)
        {
            // From
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromName, message.Sender.DisplayName);

            // Sent on
            propertyWriter.WriteProperty(SystemProperties.System.Message.DateSent, message.SentOn);

            // Subject
            propertyWriter.WriteProperty(SystemProperties.System.Subject, message.Subject);

            // Task startdate
            propertyWriter.WriteProperty(SystemProperties.System.StartDate, message.Task.StartDate);

            // Task duedate
            propertyWriter.WriteProperty(SystemProperties.System.DueDate, message.Task.DueDate);

            // Urgent
            propertyWriter.WriteProperty(SystemProperties.System.Importance, message.Importance);
            propertyWriter.WriteProperty(SystemProperties.System.ImportanceText, message.ImportanceText);

            // Status
            propertyWriter.WriteProperty(SystemProperties.System.Status, message.Task.StatusText);

            // Percentage complete
            propertyWriter.WriteProperty(SystemProperties.System.Task.CompletionStatus, message.Task.PercentageComplete);

            // Owner
            propertyWriter.WriteProperty(SystemProperties.System.Task.Owner, message.Task.Owner);

            // Categories
            propertyWriter.WriteProperty(SystemProperties.System.Category,
                message.Categories != null ? String.Join("; ", message.Categories) : null);

            // Companies
            propertyWriter.WriteProperty(SystemProperties.System.Company,
                message.Task.Companies != null ? String.Join("; ", message.Task.Companies) : null);

            // Billing information
            propertyWriter.WriteProperty(SystemProperties.System.Task.BillingInformation, message.Task.BillingInformation);

            // Mileage
            propertyWriter.WriteProperty(SystemProperties.System.MileageInformation, message.Task.Mileage);

            // Attachments
            var attachments = message.GetAttachmentNames();
            if (string.IsNullOrEmpty(attachments))
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, false);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, null);
            }
            else
            {
                propertyWriter.WriteProperty(SystemProperties.System.Message.HasAttachments, true);
                propertyWriter.WriteProperty(SystemProperties.System.Message.AttachmentNames, attachments);
            }
        }
        #endregion

        #region MapContactPropertiesToExtendedFileAttributes
        /// <summary>
        /// Maps all the filled <see cref="Storage.Task"/> properties to the corresponding extended file attributes
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="propertyWriter">The <see cref="ShellPropertyWriter"/> object</param>
        private void MapContactPropertiesToExtendedFileAttributes(Storage.Message message, ShellPropertyWriter propertyWriter)
        {
            // From
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromAddress, message.Sender.Email);
            propertyWriter.WriteProperty(SystemProperties.System.Message.FromName, message.Sender.DisplayName);

            // Sent on
            propertyWriter.WriteProperty(SystemProperties.System.Message.DateSent, message.SentOn);

            // Full name
            propertyWriter.WriteProperty(SystemProperties.System.Contact.FullName, message.Contact.DisplayName);

            // Last name
            propertyWriter.WriteProperty(SystemProperties.System.Contact.LastName, message.Contact.SurName);

            // First name
            propertyWriter.WriteProperty(SystemProperties.System.Contact.FirstName, message.Contact.GivenName);

            // Job title
            propertyWriter.WriteProperty(SystemProperties.System.Contact.JobTitle, message.Contact.Function);

            // Department
            propertyWriter.WriteProperty(SystemProperties.System.Contact.Department, message.Contact.Department);

            // Company
            propertyWriter.WriteProperty(SystemProperties.System.Company, message.Contact.Company);

            // Business address
            propertyWriter.WriteProperty(SystemProperties.System.Contact.BusinessAddress, message.Contact.WorkAddress);

            // Home address
            propertyWriter.WriteProperty(SystemProperties.System.Contact.HomeAddress, message.Contact.HomeAddress);

            // Other address
            propertyWriter.WriteProperty(SystemProperties.System.Contact.OtherAddress, message.Contact.OtherAddress);

            // Instant messaging
            propertyWriter.WriteProperty(SystemProperties.System.Contact.IMAddress, message.Contact.InstantMessagingAddress);

            // Business telephone number
            propertyWriter.WriteProperty(SystemProperties.System.Contact.BusinessTelephone, message.Contact.BusinessTelephoneNumber);

            // Assistant's telephone number
            propertyWriter.WriteProperty(SystemProperties.System.Contact.AssistantTelephone, message.Contact.AssistantTelephoneNumber);

            // Company main phone
            propertyWriter.WriteProperty(SystemProperties.System.Contact.CompanyMainTelephone, message.Contact.CompanyMainTelephoneNumber);

            // Home telephone number
            propertyWriter.WriteProperty(SystemProperties.System.Contact.HomeTelephone, message.Contact.HomeTelephoneNumber);

            // Mobile phone
            propertyWriter.WriteProperty(SystemProperties.System.Contact.MobileTelephone, message.Contact.CellularTelephoneNumber);

            // Car phone
            propertyWriter.WriteProperty(SystemProperties.System.Contact.CarTelephone, message.Contact.CarTelephoneNumber);

            // Callback
            propertyWriter.WriteProperty(SystemProperties.System.Contact.CallbackTelephone, message.Contact.CallbackTelephoneNumber);

            // Primary telephone number
            propertyWriter.WriteProperty(SystemProperties.System.Contact.PrimaryTelephone, message.Contact.PrimaryTelephoneNumber);

            // Telex
            propertyWriter.WriteProperty(SystemProperties.System.Contact.TelexNumber, message.Contact.TelexNumber);

            // TTY/TDD phone
            propertyWriter.WriteProperty(SystemProperties.System.Contact.TTYTDDTelephone, message.Contact.TextTelephone);

            // Business fax
            propertyWriter.WriteProperty(SystemProperties.System.Contact.BusinessFaxNumber, message.Contact.BusinessFaxNumber);

            // Home fax
            propertyWriter.WriteProperty(SystemProperties.System.Contact.HomeFaxNumber, message.Contact.HomeFaxNumber);

            // E-mail
            propertyWriter.WriteProperty(SystemProperties.System.Contact.EmailAddress, message.Contact.Email1EmailAddress);
            propertyWriter.WriteProperty(SystemProperties.System.Contact.EmailName, message.Contact.Email1DisplayName);

            // E-mail 2
            propertyWriter.WriteProperty(SystemProperties.System.Contact.EmailAddress2, message.Contact.Email2EmailAddress);

            // E-mail 3
            propertyWriter.WriteProperty(SystemProperties.System.Contact.EmailAddress3, message.Contact.Email3EmailAddress);

            // Birthday
            propertyWriter.WriteProperty(SystemProperties.System.Contact.Birthday, message.Contact.Birthday);

            // Anniversary
            propertyWriter.WriteProperty(SystemProperties.System.Contact.Anniversary, message.Contact.WeddingAnniversary);

            // Spouse/Partner
            propertyWriter.WriteProperty(SystemProperties.System.Contact.SpouseName, message.Contact.SpouseName);

            // Profession
            propertyWriter.WriteProperty(SystemProperties.System.Contact.Profession, message.Contact.Profession);

            // Assistant
            propertyWriter.WriteProperty(SystemProperties.System.Contact.AssistantName, message.Contact.AssistantName);

            // Web page
            propertyWriter.WriteProperty(SystemProperties.System.Contact.Webpage, message.Contact.Html);

            // Categories
            var categories = message.Categories;
            if (categories != null)
                propertyWriter.WriteProperty(SystemProperties.System.Category, String.Join("; ", String.Join("; ", categories)));
        }
        #endregion
    }
}
