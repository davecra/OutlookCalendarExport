using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookCalendarExport
{
    public class ExtendedAppointmentList : List<ExtendedAppointment>
    {
        /// <summary>
        /// Returns an instance of item based on an instance of another item that
        /// is in the provided colleciton of item
        /// </summary>
        /// <typeparam name="ExtendedAppointment"></typeparam>
        /// <param name="PobjCollection"></param>
        /// <param name="PobjPredicate"></param>
        /// <returns></returns>
        public ExtendedAppointment FindItem(ExtendedAppointment PobjOther)
        {
            try
            {
                foreach (ExtendedAppointment LobjItem in this)
                {
                    if (LobjItem.Guid == PobjOther.Guid) return LobjItem;
                }

                return default(ExtendedAppointment);
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed while trying to find the following appointment " +
                                    "in the existing list of appointments: " + PobjOther.Subject + ". " +
                                    PobjEx.Message);
            }
        }

        /// <summary>
        /// Sorts the Appointments list by date/time
        /// </summary>
        public void SortByTime()
        {
            try
            {
                for (int LintJ = this.Count - 1; LintJ > 0; LintJ--)
                {
                    for (int LintI = 0; LintI < LintJ; LintI++)
                    {
                        if (this[LintI].IsLaterThan(this[LintI + 1]))
                            exchange(LintI, LintI + 1);
                    }
                }
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Unable to sort daily appointments list. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Echsnges the two items in the array
        /// </summary>
        /// <param name="PobjM"></param>
        /// <param name="PobjN"></param>
        private void exchange(int PintA, int PintB)
        {
            // using the closing constructor
            ExtendedAppointment LobjTemp = new ExtendedAppointment(this[PintB]);

            this[PintB] = new ExtendedAppointment(this[PintA]);
            this[PintA] = new ExtendedAppointment(LobjTemp);
        }
    }
}
