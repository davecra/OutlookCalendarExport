using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookCalendarExport
{
    public class DailyAppointmentsList : List<DailyAppointments>
    {

        /// <summary>
        /// ctor
        /// </summary>
        public DailyAppointmentsList() { }

        /// <summary>
        /// ctor - initializes the class with the list of recipients, start and end date
        /// this will then loop through each day and build a list of daily event for each day and
        /// each recipient for each day and will combine the appointments that are the sdame between 
        /// recipients
        /// </summary>
        /// <param name="PobjRecipients"></param>
        /// <param name="PobjStart"></param>
        /// <param name="PobjEnd"></param>
        public bool Load(List<ExtendedRecipient> PobjRecipients, 
                         DateTime PobjStart, 
                         DateTime PobjEnd, 
                         bool PbolMeetingsOnly,
                         bool PbolExludePrivate)
        {
            try
            {
                List<DateTime> LobjDays = figureDays(PobjStart, PobjEnd);
                if (LobjDays == null)
                {
                    throw new Exception("Unable to figure days.");
                }

                Common.LoadProgress(LobjDays.Count, "Preparing...");
                foreach (DateTime LobjDay in LobjDays)
                {
                    if(!Common.IncrementProgress("Processing entires for " + LobjDay.ToLongDateString() + "..."))
                    {
                        return false; // the user cancelled
                    }
                    this.Add(new OutlookCalendarExport.DailyAppointments(PobjRecipients,
                                                                         LobjDay, 
                                                                         PbolMeetingsOnly,
                                                                         PbolExludePrivate));
                }

                Common.CloseProgress();
                return true;
            }
            catch (Exception PobjEx)
            {
                Common.CloseProgress();
                PobjEx.Log(true, "Building daily appointment list failed.");
                return false;
            }
        }

        /// <summary>
        /// Figures the days, by providing a list of dates from start ot end date
        /// </summary>
        /// <param name="PobjStart"></param>
        /// <param name="PobjEnd"></param>
        /// <returns></returns>
        private List<DateTime> figureDays(DateTime PobjStart, DateTime PobjEnd)
        {
            try
            {
                double LintTotal = (PobjEnd - PobjStart).TotalDays;
                List<DateTime> LobjDays = new List<DateTime>();
                for (int LintDay = 0; LintDay < LintTotal; LintDay++)
                {
                    LobjDays.Add(PobjStart.AddDays(LintDay));
                }
                return LobjDays;
            }
            catch(Exception PobjEx)
            {
                PobjEx.Log();
                return null;
            }
        }
    }
}
