using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;

namespace OutlookCalendarExport
{
    public class ExportToWord
    {
        public AddinSettings MobjSettings;
        public ProgressForm MobjProgress;
        private DailyAppointmentsList MobjAppointments;
        private DateTime MobjStart;
        private DateTime MobjEnd;

        public ExportToWord(AddinSettings PobjSettings)
        {
            try
            {
                MobjSettings = PobjSettings;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
            }
        }

        /// <summary>
        /// Loads all the appointments - we set thi call up
        /// </summary>
        public bool Load()
        {
            return getAllAppointments();
        }

        /// <summary>
        /// Opens Word, loads the provided template "PrintWhat" and then
        /// fills in the data based on whether the result is a:
        ///  - Day
        ///  - Work week
        ///  - Full week
        ///  - Month
        ///  The type of item to print is based on the Name of the template
        ///  provided in PrintWhat.
        /// </summary>
        public void Export()
        {
            try
            {
                // first we need to get the template type
                switch (MobjSettings.GetTemplateType())
                {
                    case AddinSettings.TemplateType.Day:
                        exportDays();
                        break;
                    case AddinSettings.TemplateType.FullWeek:
                        exportDays();
                        break;
                    case AddinSettings.TemplateType.Month:
                        // if the first day of the month is a Friday, then we start
                        // off with a count of 5, so that we start replacing the
                        // day fields on the proper day on a monthly calendar
                        DateTime LobjDate = MobjSettings.Date.GetFirstOfMonth();
                        exportDays(LobjDate.GetDayOfWeekInt());
                        break;
                    case AddinSettings.TemplateType.WorkWeek:
                        exportDays();
                        break;
                }

            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "The export to Microsoft Word failed.");
            }
        }

        /// <summary>
        /// Collect a workweeks worrth of data from the selected calendars
        /// and then open the template and replace each of the days fields
        /// </summary>
        private void exportDays(int PintStart = 1)
        {
            try
            {
                Common.LoadProgress(MobjAppointments.Count, "Preparing...");

                Word.Document LobjDoc = openDocument();
                updateHeader(LobjDoc);
                int LintCount = PintStart;
                foreach (DailyAppointments LobjDay in MobjAppointments)
                {
                    Common.IncrementProgress("Exporting day " + LobjDay.Date.ToLongDateString() + "...");
                    int LintNumber = LobjDay.Appointments.Count;

                    // update the day number - usually for monthly
                    LobjDoc.FindReplace("<<day" + LintCount.ToString() + ">>",
                                         (LintCount - PintStart + 1).GetOrdinal());

                    foreach (ExtendedAppointment LobjAppt in LobjDay.Appointments)
                    {
                        // Do we only care about shared meetings?
                        // is this a shared meetings between two or more of the
                        // recipients in this export?
                        if (LobjAppt.Recipients.Count == 1 &&
                            MobjSettings.ExportWhat == AddinSettings.ExportType.Shared)
                        { 
                            continue; // skip because we only want to see shared meetings
                        }

                        Word.Range LobjTimeRange = LobjDoc.FindReplace(
                                            "<<time" + LintCount.ToString() + ">>",
                                            LobjAppt.Start.Hour.ToString("00") + ":" +
                                            LobjAppt.Start.Minute.ToString("00") + " - " +
                                            LobjAppt.End.Hour.ToString("00") + ":" +
                                            LobjAppt.End.Minute.ToString("00"));
                        Word.Range LobjTitleRange = null;

                        string LstrSymbols = " (" + LobjAppt.Recipients.ToStringOfSymbols() + ")";
                        string LstrData = LobjAppt.Subject;

                        // see if we need to obfuscate the meeting information
                        if (MobjSettings.DisplayTimeOnly)
                        {
                            // we only going to show a generic -- Meeting -- or
                            //                                 -- Appointment --
                            if (LobjAppt.IsMeeting)
                            {
                                LstrData = " -- MEETING --";
                            }
                            else
                            {
                                LstrData = " -- APPOINTMENT --";
                            }
                        }

                        // is there is only one recipient, we do not 
                        // need symbols so lets remove them...
                        if (MobjSettings.Recipients.Count == 1)
                        {
                            // we do not need symbols
                            LstrSymbols = "";
                        }

                        // if there is no location - OR 
                        // we want to hide the lcoations
                        if (string.IsNullOrEmpty(LobjAppt.Location) ||
                            MobjSettings.ShowLocation == false)
                        {
                            // export without a location
                            LobjTitleRange = LobjDoc.FindReplace(
                                                "<<title" + LintCount.ToString() + ">>",
                                                LstrData + LstrSymbols);
                        }
                        else
                        {
                            // otherwise export with the location
                            LobjTitleRange = LobjDoc.FindReplace(
                                                "<<title" + LintCount.ToString() + ">>",
                                                LstrData + LstrSymbols +
                                                "[" + LobjAppt.Location + "]");
                        }

                        // add emphasis
                        if (LobjAppt.Recurring && MobjSettings.EmphasizeRecurring)
                        {
                            LobjTitleRange.Font.Italic = 1;
                            LobjTimeRange.Font.Italic = 1;
                        }

                        LintNumber--;
                        // if there are more appointments for the same day
                        // we want to add a new row to the inner table for that day
                        if (LintNumber > 0)
                        {
                            if (LobjTimeRange.Tables.Count > 0)
                            {
                                Word.Row LobjRow = LobjTimeRange.Rows.Add();
                                // add our find items
                                LobjRow.Cells[1].Range.Text = "<<time" + LintCount.ToString() + ">>";
                                LobjRow.Cells[2].Range.Text = "<<title" + LintCount.ToString() + ">>";
                            }
                            else
                            {
                                LobjTimeRange.InsertAfter("<<time" + LintCount.ToString() + ">>");
                                LobjTitleRange.InsertAfter("<<title" + LintCount.ToString() + ">>");
                            }
                        }
                    }
                    // increment last
                    LintCount++;
                }

                // cleanup - we remove any boilerplate items that were not used
                for (int LintX = 1; LintX <= 42; LintX++)
                {
                    LobjDoc.FindReplace("<<time" + LintX.ToString() + ">>", "");
                    LobjDoc.FindReplace("<<title" + LintX.ToString() + ">>", "");
                    LobjDoc.FindReplace("<<day" + LintX.ToString() + ">>", "");
                }

                // now delete empty tables - the cleans up unused days
                foreach (Word.Table LobjTable in LobjDoc.Tables)
                {
                    LobjTable.DeleteEmpty();
                }

                colorCode(LobjDoc); // done
                LobjDoc.Application.Visible = true;
                Common.CloseProgress();
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Failed exporting the week. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// Updates the header of the document
        /// </summary>
        /// <param name="PobjDoc"></param>
        private void updateHeader(Word.Document PobjDoc)
        {
            try
            {
                switch (MobjSettings.GetTemplateType())
                {
                    case AddinSettings.TemplateType.FullWeek:
                    case AddinSettings.TemplateType.WorkWeek:
                        PobjDoc.FindReplace("<<header>>", "Calendar for Week of " + MobjStart.Day.GetOrdinal() + " to " + MobjEnd.Day.GetOrdinal());
                        break;
                    case AddinSettings.TemplateType.Month:
                        PobjDoc.FindReplace("<<header>>", "Calendar for Month of " + MobjStart.Month.GetMonthName() + ", " + MobjEnd.Year.ToString());
                        break;
                    case AddinSettings.TemplateType.Day:
                        PobjDoc.FindReplace("<<header>>", "Day of " + MobjStart.Month.GetMonthName() + " the " + MobjStart.Day.GetOrdinal());
                        break;
                }

                if (MobjSettings.ShowHeader)
                {
                    if (MobjSettings.Recipients.Count == 1)
                    {
                        if (MobjSettings.Recipients[0].ShowName)
                        {
                            PobjDoc.FindReplace("<<name>>", "Calendar of " + 
                                                MobjSettings.Recipients[0].RecipientName);
                        }
                        else
                        {
                            PobjDoc.FindReplace("<<name>>", "Calendar of " + 
                                                MobjSettings.Recipients[0].DisplayName);
                        }
                    }
                    else
                    {
                        PobjDoc.FindReplace("<<name>>", "Calendars of " + 
                                            MobjSettings.Recipients.ToStringOfNamesWithSymbols());
                    }
                }
                else
                {
                    // delete the header row and lead paragraph
                    // this is the ONLY troubling part, we have to select
                    // the range and type a backspace
                    Word.Range LobjRange = PobjDoc.FindReplace("<<name>>", "");
                    LobjRange.Select();
                    LobjRange.Application.Selection.TypeBackspace();
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to update the header.");
            }
        }

        /// <summary>
        /// Replaces all the symbols with the proper coloring
        /// </summary>
        /// <param name="PobjDoc"></param>
        private void colorCode(Word.Document PobjDoc)
        {
            try
            {
                // now find all the symbols and color code them
                foreach (ExtendedRecipient LobjRecipient in MobjSettings.Recipients)
                {
                    Word.Find LobjFind = PobjDoc.Range().Find;
                    // Clear all previously set formatting for Find dialog box.
                    LobjFind.ClearFormatting();
                    // Clear all previously set formatting for Replace dialog box.
                    LobjFind.Replacement.ClearFormatting();
                    // Set font to Replace found font.
                    LobjFind.Text = LobjRecipient.Symbol;
                    LobjFind.Forward = true;
                    LobjFind.Wrap = Word.WdFindWrap.wdFindContinue;
                    LobjFind.Format = true;
                    LobjFind.MatchCase = true;
                    LobjFind.MatchWholeWord = false;
                    LobjFind.MatchWildcards = false;
                    LobjFind.MatchSoundsLike = false;
                    LobjFind.MatchAllWordForms = false;
                    LobjFind.Replacement.Text = LobjRecipient.Symbol;
                    LobjFind.Replacement.Font.Color = LobjRecipient.HighlightColor.FromRGBColorString().ConvertToWordColor();
                    LobjFind.Execute(Replace: Word.WdReplace.wdReplaceAll);
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to update the coloring on the symbols.");
            }
        }

        /// <summary>
        /// Opens the specified template as a new document
        /// </summary>
        /// <returns></returns>
        private Word.Document openDocument()
        {
            try
            {
                Word.Application LobjApp = new Word.Application();
                string LstrPath = "";
                if (MobjSettings.PrintWhat.StartsWith("*"))
                {
                    LstrPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Common.CUSTOMFOLDER);
                    LstrPath = Path.Combine(LstrPath, MobjSettings.PrintWhat.Replace("*", "") + "*.dotx");
                }
                else
                {
                    LstrPath = Path.Combine(Common.GetCurrentPath(), "Templates", MobjSettings.PrintWhat + ".dotx");
                }
                // open
                return LobjApp.Documents.Add(LstrPath);
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to open the template.");
                return null;
            }
        }

        private bool getAllAppointments()
        {
            try
            {
                // first we need to figure out the start and end dates
                // for all the calendar entries we will be collecting
                switch (MobjSettings.GetTemplateType())
                {
                    case AddinSettings.TemplateType.Day:
                        MobjStart = new DateTime(MobjSettings.Date.Year,
                                                 MobjSettings.Date.Month,
                                                 MobjSettings.Date.Day,
                                                 00,00,00);
                        MobjEnd = new DateTime(MobjSettings.Date.Year,
                                                 MobjSettings.Date.Month,
                                                 MobjSettings.Date.Day,
                                                 23, 59, 59); 
                        break;
                    case AddinSettings.TemplateType.FullWeek:
                        MobjStart = MobjSettings.Date.GetSunday();
                        MobjEnd = MobjStart.AddDays(6);
                        break;
                    case AddinSettings.TemplateType.WorkWeek:
                        MobjStart = MobjSettings.Date.GetMonday();
                        MobjEnd = MobjStart.AddDays(5);
                        break;
                    case AddinSettings.TemplateType.Month:
                        MobjStart = new DateTime(MobjSettings.Date.Year,
                                                 MobjSettings.Date.Month,
                                                 1);
                        MobjEnd = MobjStart.GetEndOfMonth();
                        break;
                }

                // do it
                MobjAppointments = new DailyAppointmentsList();
                bool LbolMeetingsOnly = (MobjSettings.ExportWhat != AddinSettings.ExportType.All);
                return MobjAppointments.Load(MobjSettings.Recipients, MobjStart, MobjEnd, LbolMeetingsOnly, MobjSettings.ExcludePrivate);
            }
            catch(Exception PobjEx)
            {
                PobjEx.Log();
                return false;
            }
        }
    }
}   

