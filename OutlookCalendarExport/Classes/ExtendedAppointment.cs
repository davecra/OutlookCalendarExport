using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarExport
{
    /// <summary>
    /// Public class to help build a non-object model linked
    /// appointment item so that it will perform faster
    /// </summary>
    public class ExtendedAppointment : IEquatable<ExtendedAppointment>
    {
        public DateTime Start { get; private set; }
        public DateTime End { get; private set; }
        public string Subject { get; private set; }
        public string Location { get; private set; }
        public string Guid { get; private set; }
        public bool Recurring { get; private set; }
        public bool IsMeeting { get; private set; }
        public ExtendedRecipientList Recipients { get; private set; }

        /// <summary>
        /// Created a new Extended appointment from an existing appointment
        /// </summary>
        /// <param name="PobjItem"></param>
        public ExtendedAppointment(Outlook.AppointmentItem PobjItem, ExtendedRecipient PobjRecipient)
        {
            try
            {
                Start = PobjItem.Start;
                End = PobjItem.End;
                Subject = PobjItem.Subject;
                Location = PobjItem.Location;
                Guid = PobjItem.GlobalAppointmentID;
                Recurring = PobjItem.IsRecurring;
                IsMeeting = (PobjItem.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting);
                Recipients = new ExtendedRecipientList();
                Recipients.Add(PobjRecipient);
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Unable to add appointment " + PobjItem.Subject + " from " + 
                                    PobjRecipient.RecipientName + "'s calendar. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// This is a constructor to be used for cloning an item
        /// </summary>
        /// <param name="PobjOther"></param>
        public ExtendedAppointment(ExtendedAppointment PobjItem)
        {
            try
            {
                Start = PobjItem.Start;
                End = PobjItem.End;
                Subject = PobjItem.Subject;
                Location = PobjItem.Location;
                Guid = PobjItem.Guid;
                Recurring = PobjItem.Recurring;
                IsMeeting = PobjItem.IsMeeting;
                Recipients = new ExtendedRecipientList();
                Recipients.AddRange(PobjItem.Recipients);
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Unable to clone appointment " + PobjItem.Subject + ". " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Adds a recipient to an existing appointment
        /// </summary>
        /// <param name="PobjRecipient"></param>
        public void AddRecipient(ExtendedRecipient PobjRecipient)
        {
            try
            {
                Recipients.Add(PobjRecipient);
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed to add " + PobjRecipient.RecipientName + " to a common appointment item. " +
                                    PobjEx.Message);
            }
        }

        /// <summary>
        /// Used for .Contains() to assist with easy list actions
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool Equals(ExtendedAppointment PobjOther)
        {
            try
            {
                return PobjOther.Guid == this.Guid;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
                return false;
            }
        }
    }
}
