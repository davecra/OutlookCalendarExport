using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarExport
{
    public class DailyAppointments
    {
        public ExtendedAppointmentList Appointments { get; private set; }
        public DateTime Date { get; private set; }

        public DailyAppointments(List<ExtendedRecipient> PobjRecipients, 
                                 DateTime PobjDay, 
                                 bool PbolMeetingsOnly,
                                 bool PbolExcludePrivate)
        {
            try
            {
                Date = PobjDay;
                Appointments = new ExtendedAppointmentList();
                foreach (ExtendedRecipient LobjRecipient in PobjRecipients)
                {
                    getAppointments(LobjRecipient, PobjDay, PbolMeetingsOnly, PbolExcludePrivate);
                }
                // now that we have all the appointments for 
                // the recipients for a given day we need to 
                // sort them by date time
                Appointments.SortByTime();
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Daily Appointment for recipient failed. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Adds new appointment for the specified date range
        /// </summary>
        /// <param name="PobjItem"></param>
        private void getAppointments(ExtendedRecipient PobjRecipient, 
                                     DateTime PobjDay, 
                                     bool PbolMeetingsOnly,
                                     bool PbolExcludePrivate)
        {
            try
            {
                // Start filling in the rest...
                Outlook.MAPIFolder LobjFolder = null;
                try
                {
                    LobjFolder = Common.IsRecipientValid(PobjRecipient.InteropRecipient);
                    if (LobjFolder == null)
                    {
                        throw new Exception();
                    }
                }
                catch (Exception PobjEx)
                {
                    throw new Exception("Recipient calendar folder cannot be found. " +
                                        "This might be because you have not added them as a shared calendar. " + 
                                        PobjEx.Message);
                }
                
                Outlook.Items LobjItems = LobjFolder.Items;
                if (LobjItems == null)
                {
                    throw new Exception("Unable to access recipient items. You may not have permission.");
                }

                try
                {
                    LobjItems.Sort("[Start]"); // sort the items
                }
                catch (Exception PobjEx)
                {
                    throw new Exception("Recipient calendar folder cannot be accessed or sorted. " + 
                                        "This might be because you might not have permission. " + 
                                        PobjEx.Message);
                }
                LobjItems.IncludeRecurrences = true; // be sure to include recurrences

                string LstrDay = PobjDay.ToShortDateString();
                // set the find string to today 0:00 to 23:59:59
                string LstrFind = "[Start] <= \"" + LstrDay + " 11:59 PM\"" +
                                  " AND [End] > \"" + LstrDay + " 12:00 AM\"";
                // find the first appointment for the day
                Outlook.AppointmentItem LobjAppt = LobjItems.Find(LstrFind);

                while (LobjAppt != null)
                {
                    if (LobjAppt.MeetingStatus == Outlook.OlMeetingStatus.olNonMeeting &&
                        PbolMeetingsOnly)
                    {
                        // skip - this is an appointment only
                        // and we are limiting to only meetings
                    }
                    else if(PbolExcludePrivate == true && 
                            LobjAppt.Sensitivity == Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate)
                    {
                        // skip - this is an private item
                        // and we are not includeing private items
                    }
                    else
                    {
                        ExtendedAppointment LobjNew = new ExtendedAppointment(LobjAppt, PobjRecipient);
                        if (!Appointments.Contains(LobjNew))
                        {
                            // now add the appointment
                            Appointments.Add(LobjNew);
                        }
                        else
                        {
                            Appointments.FindItem(LobjNew).AddRecipient(PobjRecipient);
                        }
                    }
                    // get the next item
                    LobjAppt = LobjItems.FindNext();
                }
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed while processing " + PobjDay.ToLongDateString() + " for " +
                                    PobjRecipient.RecipientName + ". " + PobjEx.Message);
            }
        }
    }
}
