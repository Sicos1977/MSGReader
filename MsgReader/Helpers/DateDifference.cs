using System;
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