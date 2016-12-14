using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarExport
{
    public partial class RecipientOptions : Form
    {
        private Dictionary<string, ExtendedRecipient> MobjRecipients = new Dictionary<string, ExtendedRecipient>();
        private int MintLastSelectedIndex = -1;

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="PobjRecipients"></param>
        public RecipientOptions(ExtendedRecipientList PobjRecipients)
        {
            InitializeComponent();
            foreach (ExtendedRecipient LobjItem in PobjRecipients)
            {
                listBox1.Items.Add(LobjItem.RecipientName);
                MobjRecipients.Add(LobjItem.RecipientName, LobjItem);
            }

            uiVerify();
        }

        /// <summary>
        /// Public emthod that gets a list of Extended Recipients
        /// </summary>
        /// <returns></returns>
        public ExtendedRecipientList GetRecipients()
        {
            ExtendedRecipientList LobjList = new ExtendedRecipientList();
            foreach (ExtendedRecipient LobjItem in MobjRecipients.Values)
            {
                LobjList.Add(LobjItem);
            }
            return LobjList;
        }

        private void RecipientOptions_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// The user is wanting to set the color for a user
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            ColorDialog LobjDialog = new ColorDialog();
            if (LobjDialog.ShowDialog() == DialogResult.OK)
            {
                buttonColor.BackColor = LobjDialog.Color;
            }
            buttonUpdate.Enabled = true;
        }

        /// <summary>
        /// The user wants to close the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        /// <summary>
        /// The user wants to add a new item to the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            try
            {
                Outlook.SelectNamesDialog LobjDialog;
                LobjDialog = Globals.ThisAddIn.Application.Session.GetSelectNamesDialog();
                LobjDialog.NumberOfRecipientSelectors = Microsoft.Office.Interop.Outlook.OlRecipientSelectors.olShowNone;
                if (LobjDialog.Display())
                {
                    // verify the recipient first
                    try
                    {
                        Outlook.MAPIFolder LobjFolder = Common.IsRecipientValid(LobjDialog.Recipients[1]);
                        Outlook.Items LobjItems = LobjFolder.Items;
                        LobjItems.Sort("[Start]");
                    }
                    catch 
                    {
                        MessageBox.Show("Unable to add the recipient. This might be becuase you do not have " +
                                        "permission/access to their calendar or it has not been setup as a shared " +
                                        "calendar in your Outlook Calendar.", 
                                        Common.APPNAME, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    listBox1.Items.Add(LobjDialog.Recipients[1].Name);
                    MobjRecipients.Add(LobjDialog.Recipients[1].Name, new ExtendedRecipient(LobjDialog.Recipients[1]));
                    uiVerify();
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "An error occurred when adding the name.");
            }
        }

        /// <summary>
        /// Enable/Disbale buttons based on content
        /// </summary>
        public void uiVerify()
        {
            if (listBox1.Items.Count > 0)
            {
                buttonRemove.Enabled = true;
                buttonSave.Enabled = true;
                buttonClear.Enabled = true;
            }
            else
            {
                buttonRemove.Enabled = false;
                buttonSave.Enabled = false;
                buttonClear.Enabled = false;
            }
            // cleanup
            comboBoxSymbol.Text = "";
            textBoxDisplayAs.Text = "";
            buttonColor.BackColor = SystemColors.ButtonFace;
            checkBoxShow.Checked = false;
            groupBox1.Enabled = false;
            listBox1.SelectedIndex = -1;
            buttonUpdate.Enabled = false;
        }

        /// <summary>
        /// User selected a different name in the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(listBox1.Text))
                return;
            // there are unsaved changes - ask the user first
            if (buttonUpdate.Enabled)
            {
                // change back to previous item - turn off event for this
                // then turn it back on
                listBox1.SelectedIndexChanged -= listBox1_SelectedIndexChanged;
                listBox1.SelectedIndex = MintLastSelectedIndex;
                listBox1.SelectedIndexChanged += listBox1_SelectedIndexChanged;

                // ask the user
                DialogResult LobjResult = MessageBox.Show("Do you want to save changes to the name display options?", Common.APPNAME, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (LobjResult == DialogResult.Yes)
                {
                    buttonUpdate_Click(sender, e);
                    return;
                }
                else if (LobjResult == DialogResult.Cancel)
                {
                    return; // stop
                }
            }

            // now set the last index to this item
            MintLastSelectedIndex = listBox1.SelectedIndex;
        
            // update the recipient
            ExtendedRecipient LobjRecipient = MobjRecipients[listBox1.Text];
            comboBoxSymbol.Text = LobjRecipient.Symbol;
            buttonColor.BackColor = LobjRecipient.HighlightColor.FromRGBColorString();
            textBoxDisplayAs.Text = LobjRecipient.DisplayName;
            checkBoxShow.Checked = !LobjRecipient.ShowName;
            groupBox1.Enabled = true;
            buttonUpdate.Enabled = false;
        }

        /// <summary>
        /// The user clicked the Update button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonUpdate_Click(object sender, EventArgs e)
        {
            MobjRecipients[listBox1.Text].Symbol = comboBoxSymbol.Text;
            MobjRecipients[listBox1.Text].HighlightColor = buttonColor.BackColor.ToRGBColorString();
            MobjRecipients[listBox1.Text].DisplayName = textBoxDisplayAs.Text;
            MobjRecipients[listBox1.Text].ShowName = !checkBoxShow.Checked;
            buttonUpdate.Enabled = false;
        }

        /// <summary>
        /// The user clicked the remove button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonRemove_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(listBox1.Text))
            {
                DialogResult LobjResult = MessageBox.Show("Are you sure you want to remove " + listBox1.Text + "?",
                                                          Common.APPNAME, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (LobjResult == DialogResult.Yes)
                {
                    MobjRecipients.Remove(listBox1.Text);
                    listBox1.Items.Remove(listBox1.Text);
                }
                uiVerify();
            }
        }

        /// <summary>
        /// The user selected a symbol from the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBoxSymbol_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonUpdate.Enabled = true;
        }

        /// <summary>
        /// The user types in the name box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxDisplayAs_TextChanged(object sender, EventArgs e)
        {
            buttonUpdate.Enabled = true;
        }

        /// <summary>
        /// The user chcked/unchecked the show button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBoxShow_CheckedChanged(object sender, EventArgs e)
        {
            buttonUpdate.Enabled = true;
            groupBoxDisplay.Enabled = checkBoxShow.Checked;
        }

        /// <summary>
        /// Load recipient list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonLoad_Click(object sender, EventArgs e)
        {
            try
            {
                // we load a different list from file...
                OpenFileDialog LobjDlg = new OpenFileDialog();
                LobjDlg.InitialDirectory = Common.GetUserAppDataPath(Common.APPNAME);
                LobjDlg.Filter = "Outlook Calendar Export - User List (*.oceul)|*.oceul";
                if (LobjDlg.ShowDialog() == DialogResult.OK)
                {
                    // cleanup
                    MobjRecipients = new Dictionary<string, ExtendedRecipient>();
                    listBox1.Items.Clear();

                    // load
                    ExtendedRecipientList LobjList = new ExtendedRecipientList();
                    LobjList.Import(LobjDlg.FileName);
                    foreach (ExtendedRecipient LobjItem in LobjList)
                    {
                        MobjRecipients.Add(LobjItem.RecipientName, LobjItem);
                        listBox1.Items.Add(LobjItem.RecipientName);
                    }

                    uiVerify();
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to load the recipient list.");
            }
        }

        /// <summary>
        /// The user wants to save the current list to a file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonSave_Click(object sender, EventArgs e)
        {
            try
            {
                // we save the current list to file...
                SaveFileDialog LobjDlg = new SaveFileDialog();
                LobjDlg.InitialDirectory = Common.GetUserAppDataPath(Common.APPNAME);
                LobjDlg.Filter = "Outlook Calendar Export - User List (*.oceul)|*.oceul";
                if (LobjDlg.ShowDialog() == DialogResult.OK)
                {
                    GetRecipients().Export(LobjDlg.FileName);
                    uiVerify();
                }
            }
            catch (Exception PobjEx)
            {
                PobjEx.Log(true, "Unable to load the recipient list.");
            }
        }

        /// <summary>
        /// user wants to clear the items from the list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonClear_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult LobjResult = MessageBox.Show("Are you sure you want to clear all the names?",
                                                          Common.APPNAME, MessageBoxButtons.YesNoCancel, 
                                                          MessageBoxIcon.Question);
                if (LobjResult == DialogResult.Yes)
                {
                    MobjRecipients.Clear();
                    listBox1.Items.Clear();
                    uiVerify();
                }
            }
            catch { } // ignore
        }

        /// <summary>
        /// The user wants to cancel the changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
