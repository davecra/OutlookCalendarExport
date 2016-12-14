using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarExport
{
    public partial class PrintWhatForm : Form
    {
        Outlook.Application MobjOutlook;
        public AddinSettings Settings { get; set; }

        /// <summary>
        /// Load form - set Outlook application
        /// </summary>
        public PrintWhatForm(AddinSettings PobjSettings)
        {
            try
            {
                InitializeComponent();
                MobjOutlook = Globals.ThisAddIn.Application;
                Settings = PobjSettings;
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Could not load print options form.");
                this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            }
        }

        /// <summary>
        /// User clicked Ok
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOk_Click(object sender, EventArgs e)
        {
            this.Settings.PrintWhat = listBoxTemplates.Text;
            this.Settings.ShowHeader = checkBoxHeader.Checked;
            this.Settings.Date = dtPicker.Value;
            if (radioButtonAll.Checked)
            {
                this.Settings.ExportWhat = AddinSettings.ExportType.All;
            }
            else if (radioButtonMeetings.Checked)
            {
                this.Settings.ExportWhat = AddinSettings.ExportType.Meetings;
            }
            else if (radioButtonShared.Checked)
            {
                this.Settings.ExportWhat = AddinSettings.ExportType.Shared;
            }
            this.Settings.ShowLocation = checkBoxShowLocations.Checked;
            this.Settings.EmphasizeRecurring = checkBoxRecurring.Checked;
            this.Settings.DisplayTimeOnly = checkBoxTimeOnly.Checked;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        /// <summary>
        /// USer clicked Cancel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// User changed the date
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dtPicker_ValueChanged(object sender, EventArgs e)
        {
            Settings.Date = dtPicker.Value;
            updateExportLabel();
        }

        /// <summary>
        /// On laod we connect to the server location specified
        /// in the registry or to the default folder and we then
        /// get a list of all templates in the folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PrintOptions_Load(object sender, EventArgs e)
        {
            comboBoxShow.Text = "All Templates";
            comboBoxShow.SelectedIndexChanged += ComboBoxShow_SelectedIndexChanged;
            listBoxTemplates.Items.AddRange(Common.LoadTemplates().ToArray());
            checkBoxHeader.Checked = Settings.ShowHeader;
            listBoxTemplates.Text = Settings.PrintWhat;
            switch (this.Settings.ExportWhat)
            {
                case AddinSettings.ExportType.All:
                    radioButtonAll.Checked = true;
                    break;
                case AddinSettings.ExportType.Meetings:
                    radioButtonMeetings.Checked = true;
                    break;
                case AddinSettings.ExportType.Shared:
                    radioButtonShared.Checked = true;
                    break;
            }
            checkBoxShowLocations.Checked = this.Settings.ShowLocation;
            checkBoxRecurring.Checked = this.Settings.EmphasizeRecurring;
            checkBoxTimeOnly.Checked = this.Settings.DisplayTimeOnly;

            // make sure not null
            if (Settings.Recipients != null)
            {
                txtName.Text = Settings.Recipients.ToStringOfNames();
            }
            else
            {
                txtName.Text = MobjOutlook.Session.CurrentUser.Name;
            }
            // enable / disable certain options
            if (Settings.Recipients.Count == 1)
            {
                radioButtonShared.Enabled = false;
                if (radioButtonShared.Checked)
                {
                    radioButtonMeetings.Checked = true;
                }
            }
            else
            {
                radioButtonShared.Enabled = true;
            }
            dtPicker.Value = new DateTime(DateTime.Now.Year,
                                         DateTime.Now.Month,
                                         DateTime.Now.Day,
                                         0, 0, 0, 0, DateTimeKind.Local);
            Settings.Date = dtPicker.Value;
        }

        /// <summary>
        /// User change the show to limit the template types
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBoxShow_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // load all again
                listBoxTemplates.Items.Clear();
                listBoxTemplates.Items.AddRange(Common.LoadTemplates().ToArray());
                List<string> LobjKeep = new List<string>();
                foreach (string LstrItem in listBoxTemplates.Items)
                {
                    if (LstrItem.ToUpper().StartsWith("[MONTH]") &&
                       (comboBoxShow.Text.ToUpper().Contains("MONTH") ||
                        comboBoxShow.Text.ToUpper().Contains("ALL")))
                    {
                        LobjKeep.Add(LstrItem);
                    }
                    if (LstrItem.ToUpper().StartsWith("[WORKWEEK]") &&
                       (comboBoxShow.Text.ToUpper().Contains("WORK") ||
                        comboBoxShow.Text.ToUpper().Contains("ALL")))
                    {
                        LobjKeep.Add(LstrItem);
                    }
                    if (LstrItem.ToUpper().StartsWith("[FULLWEEK]") &&
                       (comboBoxShow.Text.ToUpper().Contains("FULL") ||
                        comboBoxShow.Text.ToUpper().Contains("ALL")))
                    {
                        LobjKeep.Add(LstrItem);
                    }
                    if (LstrItem.ToUpper().StartsWith("[DAY]") &&
                       (comboBoxShow.Text.ToUpper().Contains("DAILY") ||
                        comboBoxShow.Text.ToUpper().Contains("ALL")))
                    {
                        LobjKeep.Add(LstrItem);
                    }
                }

                // now clear again and add them
                listBoxTemplates.Items.Clear();
                listBoxTemplates.Items.AddRange(LobjKeep.ToArray());
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true);
            }
        }

        /// <summary>
        /// The user wants to setup display options for multiple recipients
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonDisplayOptions_Click(object sender, EventArgs e)
        {
            try
            {
                RecipientOptions LobjForm = new RecipientOptions(Settings.Recipients);
                if (LobjForm.ShowDialog() == DialogResult.OK)
                {
                    Settings.Recipients = LobjForm.GetRecipients();
                    txtName.Text = Settings.Recipients.ToStringOfNames();
                    if (Settings.Recipients.Count == 1)
                    {
                        radioButtonShared.Enabled = false;
                        if (radioButtonShared.Checked)
                        {
                            radioButtonMeetings.Checked = true;
                        }
                    }
                    else
                    {
                        radioButtonShared.Enabled = true;
                    }
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Something happened while working with the recipients form.");
            }
        }

        

        /// <summary>
        /// Look at the item selected by the user - validate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            validate();
            // see if there is a preview in the same folder and load it
            string LstrPath = "";
            if (listBoxTemplates.Text.StartsWith("*"))
            {
                LstrPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Common.CUSTOMFOLDER);
                LstrPath = Path.Combine(LstrPath, listBoxTemplates.Text.Replace("*", "") + "*.png");
            }
            else
            {
                LstrPath = Path.Combine(Common.GetCurrentPath(), "Templates", listBoxTemplates.Text + ".png");
            }
            if (new FileInfo(LstrPath).Exists)
            {
                pictureBox1.Load(LstrPath);
                pictureBox1.AccessibleDescription = "Preview of " + listBoxTemplates.Text;
            }
            else
            {
                pictureBox1.Image = new Bitmap(1,1);
                pictureBox1.AccessibleDescription = "No preview available.";
            }
            updateExportLabel();
        }

        /// <summary>
        /// The name field has been changed validate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtName_TextChanged(object sender, EventArgs e)
        {
            validate();
            updateExportLabel();
        }

        /// <summary>
        /// Validates the form to enable to the export button
        /// </summary>
        private void validate()
        {
            if (!string.IsNullOrEmpty(txtName.Text) &&
               !string.IsNullOrEmpty(listBoxTemplates.Text))
            {
                btnOk.Enabled = true;
            }
            else
            {
                btnOk.Enabled = false;
            }
        }

        /// <summary>
        /// user checked the time only box, disable locations because
        /// we cannot show that if it is time only
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxTimeOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxTimeOnly.Checked)
            {
                checkBoxShowLocations.Enabled = false;
                checkBoxShowLocations.Checked = false;
            }
            else
            {
                checkBoxShowLocations.Enabled = true;
            }
            updateExportLabel();
        }

        /// <summary>
        /// The user clicked the about button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                FileInfo LobjFile = new FileInfo(Assembly.GetExecutingAssembly().CodeBase.Replace("file:///", "").Replace("/", "\\"));
                Assembly LobjAssm = Assembly.LoadFrom(LobjFile.FullName);
                Version LobjVersion = LobjAssm.GetName().Version;
                MessageBox.Show(Common.APPNAME + "\n\n" +
                                "This add-in is provided to give you more printing options " +
                                "in Outlook. The templates provided are fully customizable " +
                                "and new templates can be developed.\n\n" +
                                "This has been developed by:\n\n\t" +
                                " - Microsoft Premier Field Engineering\n\t" +
                                " - Date:     \t" + LobjFile.CreationTime.ToShortDateString() + "\n\t" +
                                " - Version:  \t" + LobjVersion.ToString() + "\n\n" +
                                "THIS SOFTWARE IS PROVIDED 'AS IS' AND ANY EXPRESSED OR " +
                                "IMPLIED WARRANTIES ARE DISCLAIMED. PLEASE CONTACT YOUR " +
                                "HELP DESK FOR ASSISTANCE.",
                                Common.APPNAME, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log();
            }
        }

        /// <summary>
        /// Update the lable to let the user know what they are exporting
        /// </summary>
        private void updateExportLabel()
        {
            string[] LstrNames = null;
            int LintNumber = 0;
            if (!string.IsNullOrEmpty(txtName.Text))
            {
                string LstrList = txtName.Text;
                if (txtName.Text.EndsWith(";"))
                {
                    // remove last char
                    LstrList = LstrList.Substring(0,LstrList.Length - 1);
                }
                LstrNames = LstrList.Split(';');
                LintNumber = LstrNames.Length;
            }

            // start...
            string LstrResult = "Exporting the calendar";
            if(LintNumber > 1)
            {
                LstrResult += "s of ";
            }
            else
            {
                LstrResult += " of ";
            }

            // build a list of names 
            if (LintNumber == 0)
            {
                LstrResult += "[no names selected yet]";
            }
            else
            {
                // list all the names
                foreach (string LstrName in LstrNames)
                {
                    if (LintNumber > 2)
                    {
                        LstrResult += LstrName + ", ";
                    }
                    else if (LintNumber == 2)
                    {
                        if (LstrNames.Length == 2)
                        {
                            LstrResult += LstrName + " and ";
                        }
                        else
                        {
                            LstrResult += LstrName + ", and ";
                        }
                    }
                    else
                    {
                        LstrResult += LstrName;
                    }
                    LintNumber--;
                }
            }

            // next, list the dates
            if (listBoxTemplates.Text.ToUpper().StartsWith("[DAY]"))
            {
                LstrResult += " for the day of " + dtPicker.Value.ToLongDateString(); 
            }
            else if (listBoxTemplates.Text.ToUpper().StartsWith("[FULLWEEK]"))
            {
                LstrResult += " for the week of Sunday (" +
                              dtPicker.Value.GetSunday().ToShortDateString() +
                              ") to Saturday, (" +
                              dtPicker.Value.GetSunday().AddDays(6.99).ToShortDateString() + ")";
            }
            else if (listBoxTemplates.Text.ToUpper().StartsWith("[WORKWEEK]"))
            {
                LstrResult += " for the week of Monday (" +
                              dtPicker.Value.GetMonday().ToShortDateString() +
                              ") to Friday (" +
                              dtPicker.Value.GetMonday().AddDays(4.99).ToShortDateString() + ")";
            }
            else if (listBoxTemplates.Text.ToUpper().StartsWith("[MONTH]"))
            {
                LstrResult += " for the month of " + dtPicker.Value.GetMonthName() +
                              " " + dtPicker.Value.Year.ToString();
            }
            else
            {
                LstrResult += " [no template selected yet]";
            }

            // next list what is being exported
            if (radioButtonAll.Checked)
            {
                LstrResult += ", showing all appointments and meetings";
            }
            else if (radioButtonMeetings.Checked)
            {
                LstrResult += ", showing only meetings";
            }
            else
            {
                LstrResult += ", showing meetings attended together";
            }

            if (checkBoxShowLocations.Checked)
            {
                LstrResult += ", meeting locations will be included";
            }

            if (checkBoxTimeOnly.Checked)
            {
                LstrResult += ", only the times will be shown (subjects will be generic)";
            }

            if (checkBoxHeader.Checked)
            {
                LstrResult += ", names will be included in the header";
            }

            if (checkBoxRecurring.Checked)
            {
                LstrResult += ", recurring appointments will be in italics";
            }

            // replace final comma with a comma and an "and"
            int LintLast = LstrResult.LastIndexOf(",");
            LstrResult = LstrResult.ReplaceLastInstanceOf(",", ", and");
            LstrResult += ".";

            // now set it
            labelWhat.Text = LstrResult;
        }

        /// <summary>
        /// User clicked the radio button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButtonAll_CheckedChanged(object sender, EventArgs e)
        {
            updateExportLabel();
        }

        /// <summary>
        /// User clicked the radio button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButtonMeetings_CheckedChanged(object sender, EventArgs e)
        {
            updateExportLabel();
        }

        /// <summary>
        /// User clicked the radio button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButtonShared_CheckedChanged(object sender, EventArgs e)
        {
            updateExportLabel();
        }

        /// <summary>
        /// User clicked the checkbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxShowLocations_CheckedChanged(object sender, EventArgs e)
        {
            updateExportLabel();
        }

        /// <summary>
        /// User clicked the checkbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxHeader_CheckedChanged(object sender, EventArgs e)
        {
            updateExportLabel();
        }

        /// <summary>
        /// User clicked the checkbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxRecurring_CheckedChanged(object sender, EventArgs e)
        {
            updateExportLabel();
        }
    }
}
