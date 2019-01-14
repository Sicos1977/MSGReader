//
// DateDifference.cs
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
using MsgReader.Localization;

namespace MsgReader.Helpers
{
    /// <summary>
    /// This class can be used to calculate the differences between 2 dates in years, months, weeks, days, hours and seconds
    /// </summary>
    internal class DateDifference
    {
        #region Properties
        public int Years { get; private set; }
        public int Months { get; private set; }
        public int Weeks { get; private set; }
        public int Days { get; private set; }
        public int Hours { get; private set; }
        public int Minutes { get; private set; }
        public int Seconds { get; private set; }
        #endregion

        #region ToString
        /// <summary>
        /// Returns the years, days, months, weeks, hours, minutes and seconds as a string.
        /// E.g. 1 year, 2 months, 3 weeks, 4 days, 5 hours, 6 minutes
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            var result = string.Empty;

            if (Years > 0)
            {
                if (Years == 1)
                    result += Years + " " + LanguageConsts.DateDifferenceYearText;
                else
                    result += Years + " " + LanguageConsts.DateDifferenceYearsText;
            }

            if (Months > 0)
            {
                if (!string.IsNullOrEmpty(result))
                    result += ", ";

                if (Months == 1)
                    result += Months + " " + LanguageConsts.DateDifferenceMonthText;
                else
                    result += Months + " " + LanguageConsts.DateDifferenceMonthsText;
            }

            if (Weeks > 0)
            {
                if (!string.IsNullOrEmpty(result))
                    result += ", ";

                if (Weeks == 1)
                    result += Weeks + " " + LanguageConsts.DateDifferenceWeekText;
                else
                    result += Weeks + " " + LanguageConsts.DateDifferenceWeeksText;
            }

            if (Days > 0)
            {
                if (!string.IsNullOrEmpty(result))
                    result += ", ";

                if (Days == 1)
                    result += Days + " " + LanguageConsts.DateDifferenceDayText;
                else
                    result += Days + " " + LanguageConsts.DateDifferenceDaysText;
            }

            if (Hours > 0)
            {
                if (!string.IsNullOrEmpty(result))
                    result += ", ";

                if (Hours == 1)
                    result += Hours + " " + LanguageConsts.DateDifferenceHourText;
                else
                    result += Hours + " " + LanguageConsts.DateDifferenceHoursText;
            }

            if (Minutes > 0)
            {
                if (!string.IsNullOrEmpty(result))
                    result += ", ";

                if (Hours == 1)
                    result += Minutes + " " + LanguageConsts.DateDifferenceMinuteText;
                else
                    result += Minutes + " " + LanguageConsts.DateDifferenceMinutesText;
            }

            if (Seconds > 0)
            {
                if (!string.IsNullOrEmpty(result))
                    result += ", ";

                if (Seconds == 1)
                    result += Seconds + " " + LanguageConsts.DateDifferenceSecondText;
                else
                    result += Seconds + " " + LanguageConsts.DateDifferenceSecondsText;
            }

            // Default result for 0 hours when there is no difference
            if (string.IsNullOrEmpty(result))
                result = "0 " + LanguageConsts.DateDifferenceHourText;

            return result;
        }
        #endregion

        #region Difference
        /// <summary>
        /// Calculates the difference between 2 dates in years, months, weeks, days, hours and seconds
        /// </summary>
        /// <param name="dateTime1"></param>
        /// <param name="dateTime2"></param>
        /// <returns></returns>
        public static DateDifference Difference(DateTime dateTime1, DateTime dateTime2)
        {
            if (dateTime1 > dateTime2)
            {
                var dtTemp = dateTime1;
                dateTime1 = dateTime2;
                dateTime2 = dtTemp;
            }

            var dateDiff = new DateDifference {Years = dateTime2.Year - dateTime1.Year};
            if (dateDiff.Years > 0)
                if (dateTime2.Month < dateTime1.Month)
                    dateDiff.Years--;
                else if (dateTime2.Month == dateTime1.Month)
                    if (dateTime2.Day < dateTime1.Day)
                        dateDiff.Years--;
                    else if (dateTime2.Day == dateTime1.Day)
                        if (dateTime2.Hour < dateTime1.Hour)
                            dateDiff.Years--;
                        else if (dateTime2.Hour == dateTime1.Hour)
                            if (dateTime2.Minute < dateTime1.Minute)
                                dateDiff.Years--;
                            else if (dateTime2.Minute == dateTime1.Minute)
                                if (dateTime2.Second < dateTime1.Second)
                                    dateDiff.Years--;

            dateDiff.Months = dateTime2.Month - dateTime1.Month;
            if (dateTime2.Month < dateTime1.Month)
                dateDiff.Months = 12 - dateTime1.Month + dateTime2.Month;
            
            if (dateDiff.Months > 0)
                if (dateTime2.Day < dateTime1.Day)
                    dateDiff.Months--;
                else if (dateTime2.Day == dateTime1.Day)
                    if (dateTime2.Hour < dateTime1.Hour)
                        dateDiff.Months--;
                    else if (dateTime2.Hour == dateTime1.Hour)
                        if (dateTime2.Minute < dateTime1.Minute)
                            dateDiff.Months--;
                        else if (dateTime2.Minute == dateTime1.Minute)
                            if (dateTime2.Second < dateTime1.Second)
                                dateDiff.Months--;

            dateDiff.Days = dateTime2.Day - dateTime1.Day;
            if (dateTime2.Day < dateTime1.Day)
                dateDiff.Days = DateTime.DaysInMonth(dateTime1.Year, dateTime1.Month) - dateTime1.Day + dateTime2.Day;
            
            if (dateDiff.Days > 0)
                if (dateTime2.Hour < dateTime1.Hour)
                    dateDiff.Days--;
                else if (dateTime2.Hour == dateTime1.Hour)
                    if (dateTime2.Minute < dateTime1.Minute)
                        dateDiff.Days--;
                    else if (dateTime2.Minute == dateTime1.Minute)
                        if (dateTime2.Second < dateTime1.Second)
                            dateDiff.Days--;

            dateDiff.Weeks = dateDiff.Days/7;
            dateDiff.Days = dateDiff.Days%7;

            dateDiff.Hours = dateTime2.Hour - dateTime1.Hour;
            if (dateTime2.Hour < dateTime1.Hour)
                dateDiff.Hours = 24 - dateTime1.Hour + dateTime2.Hour;
        
            if (dateDiff.Hours > 0)
                if (dateTime2.Minute < dateTime1.Minute)
                    dateDiff.Hours--;
                else if (dateTime2.Minute == dateTime1.Minute)
                    if (dateTime2.Second < dateTime1.Second)
                        dateDiff.Hours--;

            dateDiff.Minutes = dateTime2.Minute - dateTime1.Minute;
            if (dateTime2.Minute < dateTime1.Minute)
                dateDiff.Minutes = 60 - dateTime1.Minute + dateTime2.Minute;
        
            if (dateDiff.Minutes > 0)
                if (dateTime2.Second < dateTime1.Second)
                    dateDiff.Minutes--;
        
            dateDiff.Seconds = dateTime2.Second - dateTime1.Second;
            
            if (dateTime2.Second < dateTime1.Second)
                dateDiff.Seconds = 60 - dateTime1.Second + dateTime2.Second;
            
            return dateDiff;
        }
        #endregion
    }
}