using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OutlookCalendarExport
{
    public partial class RibbonExport
    {
        Outlook.Application MobjOutlook;
        ProgressForm MobjProgress = null;

        /// <summary>
        /// Get reference to Outlook from Addin
        /// </summary>
        /// <param name="PobjSender"></param>
        /// <param name="PobjEventArgs"></param>
        private void Ribbon1_Load(object PobjSender, RibbonUIEventArgs PobjEventArgs)
        {
            MobjOutlook = Globals.ThisAddIn.Application;
        }

        /// <summary>
        /// Ask the user for the date they want to print with Today
        /// Selected by default. Then open Word, set the sheet size to 5x3, 
        /// insert a table 2 columns a merged header (2 rows) and then
        /// proceed to fill it with the information from the calendar.
        /// Then open the Print Dialog for Word...
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonExport_Click(object PobjSender, RibbonControlEventArgs PobjEventArgs)
        {
            PrintWhatForm LobjDlg = null;
            try
            {
                AddinSettings LobjSettings = new AddinSettings();
                LobjSettings.LoadSettings();

                // show dialog
                LobjDlg = new PrintWhatForm(LobjSettings);
                if (LobjDlg.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                    return; // done

                LobjSettings = LobjDlg.Settings;

                // save the settings
                LobjSettings.SaveSettings();

                // DO IT
                ExportToWord LobjExport = new ExportToWord(LobjSettings);
                if (LobjExport.Load()) 
                {
                    // load of items successful - export
                    LobjExport.Export();
                    // done
                    MessageBox.Show("Completed!", Common.APPNAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "There was an error exporting to Word.");
            }
            finally
            {
                if (MobjProgress != null)
                {
                    MobjProgress.Close();
                    MobjProgress = null;
                }
            }
        }
    }
}
