using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Threading;
using System.Diagnostics;

namespace OutlookCalendarExport
{ 
    public static class Common
    {
        public const float POINTS = 72.0f;
        public const string APPNAME = "Export Outlook Calendar to Word Add-in";
        public const string REGPATH = "Software\\Microsoft\\OutlookCalendarExport";
        public const string CUSTOMFOLDER = "Outlook Custom Calendar Templates";

        /// <summary>
        /// Returns a full path to the specified folder in the AppData
        /// </summary>
        /// <param name="PstrFolder"></param>
        /// <returns></returns>
        public static string GetUserAppDataPath(string PstrFolder)
        {
            try
            {
                string LstrPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                LstrPath = Path.Combine(LstrPath, PstrFolder);
                if (!new DirectoryInfo(LstrPath).Exists)
                {
                    new DirectoryInfo(LstrPath).Create();
                }
                return LstrPath;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to get AppData path.");
                return "";
            }
        }

        /// <summary>
        /// Enum parser helper
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enum"></param>
        /// <returns></returns>
        public static T GetEnumFromName<T>(this object @enum)
        {
            return (T)Enum.Parse(typeof(T), @enum.ToString());
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Checks - recursively - if the given table is empty
        /// </summary>
        /// <param name="PobjTable"></param>
        /// <returns></returns>
        public static void DeleteEmpty(this Word.Table PobjTable)
        {
            try
            {
                foreach (Word.Table PobjItem in PobjTable.Tables)
                {
                    PobjItem.DeleteEmpty();
                }
                // figure to toal number of characters in this table if it were empty
                // each blank cell has two (2) characters 
                int LintTotal = PobjTable.Range.Cells.Count * 2;
                int LintCount = 0;
                foreach (Word.Cell LobjCell in PobjTable.Range.Cells)
                {
                    LintCount += LobjCell.Range.Text.Length;
                }

                // now we see if it is empty and if it is then we delete it
                if (LintTotal >= LintCount)
                {
                    PobjTable.Delete();
                }
                else
                {
                    PobjTable.DeleteRows();
                }
            }
            catch { }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// loops through a table and deletes all empty rows
        /// </summary>
        /// <param name="PobjTable"></param>
        public static void DeleteRows(this Word.Table PobjTable)
        {
            // now delete empty rows too
            foreach (Word.Row LobjRow in PobjTable.Rows)
            {
                int LintTotal = LobjRow.Range.Cells.Count * 2;
                int LintCount = 0;
                foreach (Word.Cell LobjCell in LobjRow.Cells)
                {
                    LintCount += LobjCell.Range.Text.Length;
                }
                // if there are not characters - delete
                if (LintTotal >= LintCount)
                {
                    LobjRow.Delete();
                }
            }
        }

        /// <summary>
        /// Return an integer representing the day of the week
        ///  Sunday = 1
        /// </summary>
        /// <param name="PobjDate"></param>
        /// <returns></returns>
        public static int GetDayOfWeekInt(this DateTime PobjDate)
        {
            switch (PobjDate.DayOfWeek)
            {
                case DayOfWeek.Sunday: return 1;
                case DayOfWeek.Monday: return 2;
                case DayOfWeek.Tuesday: return 3;
                case DayOfWeek.Wednesday: return 4;
                case DayOfWeek.Thursday: return 5;
                case DayOfWeek.Friday: return 6;
                case DayOfWeek.Saturday: return 7;
                default: return 1; // sunday
            }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Log the exception to temp folder and notify the user if the Prompt is enabled
        /// </summary>
        /// <param name="PobjEx"></param>
        /// <param name="PbolPrompt"></param>
        /// <param name="PstrMessage"></param>
        public static void Log(this Exception PobjEx, bool PbolPrompt = false, string PstrMessage = "")
        {
            try
            {
                if (PbolPrompt && !string.IsNullOrEmpty(PstrMessage))
                {
                    MessageBox.Show(PobjEx.Message + "\n\n" + PstrMessage, Common.APPNAME,
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (PbolPrompt)
                {
                    MessageBox.Show(PobjEx.Message, Common.APPNAME,
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // log it
                string LstrFn = Path.Combine(Path.GetTempPath(),
                                             Common.APPNAME + "_" + DateTime.Now.ConvertToRaw() + ".log");
                StreamWriter LobjSw = new StreamWriter(LstrFn);
                LobjSw.WriteLine(DateTime.Now.ToString() + "\n" +
                                 PobjEx.ToString() + "\n" +
                                 PstrMessage);
                LobjSw.Close();
            }
            catch { } // ignore
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Converts a list of extended recipients to a string of names
        /// </summary>
        /// <param name="PobjList"></param>
        /// <returns></returns>
        public static string ToStringOfNames(this List<ExtendedRecipient> PobjList)
        {
            try
            {
                string LstrResult = "";
                PobjList.ForEach(delegate (ExtendedRecipient PobjRecipient)
                {
                    LstrResult += PobjRecipient.RecipientName + ";";
                });
                return LstrResult;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Converts a list of extended recipients to a string of names with symbols
        /// </summary>
        /// <param name="PobjList"></param>
        /// <returns></returns>
        public static string ToStringOfNamesWithSymbols(this List<ExtendedRecipient> PobjList)
        {
            try
            {
                string LstrResult = "";
                PobjList.ForEach(delegate (ExtendedRecipient PobjRecipient)
                {
                    if(PobjRecipient.ShowName)
                    {
                        LstrResult += PobjRecipient.RecipientName + " (" + PobjRecipient.Symbol + "), ";
                    }
                    else
                    {
                        LstrResult += PobjRecipient.DisplayName + " (" + PobjRecipient.Symbol + "), ";
                    }
                });
                return LstrResult;
            }
            catch
            {
                return "";
            }
        }

        public static string ToStringOfSymbols(this List<ExtendedRecipient> PobjList)
        {
            try
            {
                string LstrResult = "";
                PobjList.ForEach(delegate (ExtendedRecipient PobjRecipient)
                {
                    LstrResult += PobjRecipient.Symbol + " ";
                });
                return LstrResult.Trim();
            }
            catch
            {
                return "";
            }
        }

        public static Word.Range FindReplace(this Word.Document PobjDoc, string PstrFind, string PstrReplace)
        {
            try
            {
                Word.Range LobjRange = PobjDoc.Range();
                LobjRange.Find.Execute2007(FindText: PstrFind, ReplaceWith: PstrReplace);
                return LobjRange;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Converts a System Color object to a Word Color Object
        /// </summary>
        /// <param name="PobjColor"></param>
        /// <returns></returns>
        public static Word.WdColor ConvertToWordColor(this Color PobjColor)
        {
            try
            {
                return (Microsoft.Office.Interop.Word.WdColor)(PobjColor.R + 0x100 * PobjColor.G + 0x10000 * PobjColor.B);
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
                return Word.WdColor.wdColorBlack;
            }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Takes a Recipients Object list and converts it to a 
        /// List of Outlook Recipient objects
        /// </summary>
        /// <param name="PobjRecipients"></param>
        /// <returns></returns>
        public static ExtendedRecipientList ToListOfExtendedRecipient(this Outlook.Recipients PobjRecipients)
        {
            try
            {
                ExtendedRecipientList LobjResult = new ExtendedRecipientList();
                foreach (Outlook.Recipient LobjItem in PobjRecipients)
                {
                    LobjResult.Add(new ExtendedRecipient(LobjItem));
                }
                return LobjResult;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Load the templates from the current path
        /// </summary>
        /// <returns></returns>
        public static List<string> LoadTemplates()
        {
            List<string> LstrResult = new List<string>();
            try
            {
                string LstrPath = Path.Combine(GetCurrentPath(), "Templates");
                foreach (FileInfo LobjFile in new DirectoryInfo(LstrPath).GetFiles())
                {
                    // add only Word template files and not tempa files
                    if(LobjFile.Extension.ToLower() == ".dotx" && !LobjFile.Name.StartsWith("~"))
                        LstrResult.Add(LobjFile.Name.ToUpper().Replace(".DOTX",""));
                }
                // custom templates
                LstrPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Common.CUSTOMFOLDER);
                if (!new DirectoryInfo(LstrPath).Exists)
                {
                    new DirectoryInfo(LstrPath).Create();
                }
                foreach (FileInfo LobjFile in new DirectoryInfo(LstrPath).GetFiles())
                {
                    // add only Word template files and not tempa files
                    if (LobjFile.Extension.ToLower() == ".dotx" && !LobjFile.Name.StartsWith("~"))
                        LstrResult.Add("*" + LobjFile.Name.ToUpper().Replace(".DOTX", ""));
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to load templates.");
            }
            return LstrResult;
        }

        /// <summary>
        /// Gets the path to the current DLL install location
        /// </summary>
        /// <returns></returns>
        public static string GetCurrentPath()
        {
            try
            {
                return Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase).Replace("file:\\", "").Replace("/", "\\");
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
                return "";
            }
        }

        /// <summary>
        /// Converts a color to a 9 digit string, for example Red will become:
        ///         255000000
        /// </summary>
        /// <param name="PobjColor"></param>
        /// <returns></returns>
        public static string ToRGBColorString(this Color PobjColor)
        {
            return PobjColor.R.ToString("000") + PobjColor.G.ToString("000") + PobjColor.B.ToString("000");
        }

        /// <summary>
        /// Takes a string with 9 characters 00000000 and convert it to RGB color
        /// </summary>
        /// <param name="PstrColor"></param>
        /// <returns></returns>
        public static Color FromRGBColorString(this string PstrColor)
        {
            if (PstrColor.Length == 9)
            {
                return Color.FromArgb(int.Parse(PstrColor.Substring(0, 3)),
                                      int.Parse(PstrColor.Substring(3, 3)),
                                      int.Parse(PstrColor.Substring(6, 3)));
            }
            else
            {
                return Color.Black;
            }
        }

        /// <summary>
        /// Take a number and return a string with the name of the month
        /// </summary>
        /// <param name="PintValue"></param>
        /// <returns></returns>
        public static string GetMonthName(this int PintValue)
        {
            switch (PintValue)
            {
                case 1: return "January";
                case 2: return "February";
                case 3: return "March";
                case 4: return "April";
                case 5: return "May";
                case 6: return "June";
                case 7: return "July";
                case 8: return "August";
                case 9: return "September";
                case 10: return "October";
                case 11: return "November";
                case 12: return "December";
                default: return "";
            }
        }

        /// <summary>
        /// Given a string to find, this function will remove the very last
        /// instance of the string and replace it with the provided string
        /// </summary>
        /// <param name="PstrValue"></param>
        /// <param name="PstrFind"></param>
        /// <param name="PstrReplace"></param>
        /// <returns></returns>
        public static string ReplaceLastInstanceOf(this string PstrValue, string PstrFind, string PstrReplace)
        {
            try
            {
                int LintPos = PstrValue.LastIndexOf(PstrFind);
                string LstrFront = PstrValue.Substring(0, LintPos);
                string LstrEnd = PstrValue.Substring(LintPos + PstrFind.Length);
                string LstrReturn = LstrFront + PstrReplace + LstrEnd;
                return LstrReturn;
            }
            catch
            {
                return ""; // failed
            }
        }

        /// <summary>
        /// Given a date, this function returns the month name
        /// </summary>
        /// <param name="PobjValue"></param>
        /// <returns></returns>
        public static string GetMonthName(this DateTime PobjValue)
        {
            return PobjValue.Month.GetMonthName();
        }

        /// <summary>
        /// Take a number and retusn a string with its ordinal name: 1st, 2nd, 3rd...
        /// </summary>
        /// <param name="PintValue"></param>
        /// <returns></returns>
        public static string GetOrdinal(this int PintValue)
        {
            string LstrVal = PintValue.ToString();
            string LstrOrd = LstrVal.Substring(LstrVal.Length - 1);

            switch (int.Parse(LstrOrd))
            {
                case 1:
                    if (PintValue == 11)
                    {
                        return PintValue.ToString() + "th";
                    }
                    else
                    {
                        return PintValue.ToString() + "st";
                    }
                case 2:
                    if (PintValue == 12)
                    {
                        return PintValue.ToString() + "th";
                    }
                    else
                    {
                        return PintValue.ToString() + "nd";
                    }
                case 3:
                    if (PintValue == 13)
                    {
                        return PintValue.ToString() + "th";
                    }
                    else
                    {
                        return PintValue.ToString() + "rd";
                    }
                default:
                    return PintValue.ToString() + "th"; 
            }
        }

        /// <summary>
        /// Get the recipient dialog and allow the user to chose a recipient
        /// </summary>
        /// <param name="PstrRecipient"></param>
        public static ExtendedRecipientList GetRecipients(string PstrRecipient = "")
        {
            try
            {
                Outlook.SelectNamesDialog LobjSnd = Globals.ThisAddIn.Application.Session.GetSelectNamesDialog();
                if (PstrRecipient != string.Empty)
                    LobjSnd.Recipients.Add(PstrRecipient);
                LobjSnd.NumberOfRecipientSelectors = Outlook.OlRecipientSelectors.olShowTo;
                LobjSnd.AllowMultipleSelection = false;
                LobjSnd.Display();
                if (!LobjSnd.Recipients.ResolveAll())
                {
                    return null;
                }
                else
                {
                    return LobjSnd.Recipients.ToListOfExtendedRecipient();
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "There was an error selecting the recipient.");
                return null;
            }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Returns a date that has no spaces and is all apha characters
        /// </summary>
        /// <param name="PobjDate"></param>
        /// <returns></returns>
        public static string ConvertToRaw(this DateTime PobjDate)
        {
            try
            {
                string LstrResult = DateTime.Now.ToString();
                LstrResult = LstrResult.Replace("\\", "")
                                       .Replace("/", "")
                                       .Replace(" ", "");
                return LstrResult;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Returns the first day of the full week (Sunday)
        /// </summary>
        /// <param name="PobjDate"></param>
        /// <returns></returns>
        public static DateTime GetSunday(this DateTime PobjDate)
        {
            try
            {
                double LintDoW = (double)PobjDate.DayOfWeek;
                return PobjDate.AddDays(-1 * LintDoW);
            }
            catch(Exception PobjEx)
            {
                PobjEx.Log();
                return PobjDate;
            }
        }

        /// <summary>
        /// Returns the first day of the work week (Monday)
        /// </summary>
        /// <param name="PobjDate"></param>
        /// <returns></returns>
        public static DateTime GetMonday(this DateTime PobjDate)
        {
            try
            {
                double LintDoW = (double)PobjDate.DayOfWeek;
                return PobjDate.AddDays((-1 * LintDoW) + 1);
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
                return PobjDate;
            }
        }

        /// <summary>
        /// Returns the first day of the month
        /// </summary>
        /// <param name="PobjDate"></param>
        /// <returns></returns>
        public static DateTime GetFirstOfMonth(this DateTime PobjDate)
        {
            try
            {
                return new DateTime(PobjDate.Year, PobjDate.Month, 1);
            }
            catch(Exception PobjEx)
            {
                PobjEx.Log();
                return PobjDate;
            }
        }

        /// <summary>
        /// Returns the last day of the month
        /// </summary>
        /// <param name="PobjDate"></param>
        /// <returns></returns>
        public static DateTime GetEndOfMonth(this DateTime PobjDate)
        {
            try
            {
                DateTime LobjNextMonth = new DateTime(PobjDate.Year, PobjDate.Month + 1, 1);
                return LobjNextMonth.AddDays(-1);
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
                return PobjDate;
            }
        }

        /// <summary>
        /// Retuns the calendar folder for the recipient if they are visible
        /// in the Outlook Explorer / added
        /// </summary>
        /// <param name="PobjRecipient"></param>
        /// <returns></returns>
        public static Outlook.MAPIFolder IsRecipientValid(Outlook.Recipient PobjRecipient)
        {
            // see if the selected user's calendar is available in Outlook
            // and then switch to it. Otherwise, give an error
            Outlook.MAPIFolder LobjFolder = null;
            try
            {
                Outlook.Application LobjApp = Globals.ThisAddIn.Application;
                Outlook.NameSpace LobjNs = LobjApp.GetNamespace("MAPI");
                LobjFolder = LobjNs.GetSharedDefaultFolder(
                                PobjRecipient, Outlook.OlDefaultFolders.olFolderCalendar)
                                as Outlook.MAPIFolder;
                return LobjFolder;
            }
            catch (Exception PobjEx)
            {
                throw new Exception("Could not open users's calendar. " + PobjEx.Message);
            }
        }

        /// <summary>
        /// EXTENSION METHOD
        /// Compares whether this ExtendedAppointment occurs after (greater than) the
        /// appointment that it is being compared to
        /// </summary>
        /// <param name="PobjThis"></param>
        /// <param name="PobjThat"></param>
        /// <returns></returns>
        public static bool IsLaterThan(this ExtendedAppointment PobjThis, ExtendedAppointment PobjThat)
        {
            if (PobjThis.Start > PobjThat.Start)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Hash the date
        /// </summary>
        /// <param name="PobjDate"></param>
        /// <returns></returns>
        public static double HashDate(DateTime PobjDate)
        {
            TimeSpan LobjSpan = PobjDate - DateTime.Parse("1/1/2013");
            return LobjSpan.TotalMinutes;
        }

        /// <summary>
        /// Formatted time string for the output
        /// </summary>
        /// <param name="PobjDateStart"></param>
        /// <param name="PobjDateEnd"></param>
        /// <returns></returns>
        public static string GetTimeString(DateTime PobjDateStart, DateTime PobjDateEnd)
        {
            return PobjDateStart.Hour.ToString("00") + ":" + PobjDateStart.Minute.ToString("00") + " - " +
                   PobjDateEnd.Hour.ToString("00") + ":" + PobjDateEnd.Minute.ToString("00");
        }

        /// <summary>
        /// Increment the progress bar and check for the user
        /// pressing cancel on the dialog
        /// </summary>
        /// <param name="PbolIncrement"></param>
        /// <returns></returns>
        public static bool incrementProgress(this Form LobjForm, bool PbolIncrement = true, bool PbolStall = false)
        {
            if (PbolIncrement && LobjForm.Controls.ContainsKey("ProgressBar"))
                LobjForm.Controls["ProgressBar"].GetType().InvokeMember("Increment", System.Reflection.BindingFlags.InvokeMethod, null, LobjForm, new object[] { 10 });
            LobjForm.Refresh();
            System.Windows.Forms.Application.DoEvents();
            if (!LobjForm.Visible)
            {
                LobjForm.Close();
                return false; // exit-stop
            }
            else
            {
                // delay to allow finish
                if (PbolStall)
                {
                    DateTime dtStall = DateTime.Now.AddSeconds(1);
                    while (dtStall > DateTime.Now)
                        System.Windows.Forms.Application.DoEvents();
                }
                return true;
            }
        }

        public static ProgressForm GobjProgress;

        /// <summary>
        /// Load and initialize the progress form or set it to
        /// a new value to reset it to zero
        /// </summary>
        /// <param name="PintMax"></param>
        /// <param name="PstrValue"></param>
        public static void LoadProgress(int PintMax, string PstrValue)
        {
            try
            {
                IntPtr LintHwnd = Process.GetProcessesByName("Outlook")[0].MainWindowHandle;
                if (GobjProgress == null)
                {
                    GobjProgress = new ProgressForm(PintMax, PstrValue);
                    GobjProgress.Show(new ArbitraryWindow(LintHwnd));
                }
                else
                {
                    GobjProgress.SetLabel(PstrValue);
                    GobjProgress.SetProgressBarMax(PintMax);
                    if (!GobjProgress.Visible)
                    {
                        GobjProgress.Show(new ArbitraryWindow(LintHwnd));
                    }
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to load the progress form.");
            }
        }

        /// <summary>
        /// Increment the progress form by one
        /// </summary>
        /// <param name="PstrValue"></param>
        /// <returns></returns>
        public static bool IncrementProgress(string PstrValue = "")
        {
            try
            {
                Application.DoEvents();
                if (GobjProgress != null)
                {
                    // user closed - cancelled
                    if (GobjProgress.UserCancelledMe)
                    {
                        CloseProgress();
                        Application.DoEvents();
                        return false;
                    }

                    if (!string.IsNullOrEmpty(PstrValue))
                    {
                        GobjProgress.SetLabel(PstrValue);
                    }

                    // increment
                    GobjProgress.Increment();

                    // refresh
                    GobjProgress.Refresh();

                    // ok
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return true; // not massive - ignore
            } 
        }

        /// <summary>
        /// Closes the progress form
        /// </summary>
        public static void CloseProgress()
        {
            try
            {
                if (GobjProgress != null)
                    GobjProgress.Close();
                GobjProgress = null;
            }
            catch { } // ignore
        }
    }

    class ArbitraryWindow : IWin32Window
    {
        public ArbitraryWindow(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; private set; }
    }
}
