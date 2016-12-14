using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookCalendarExport
{
    public partial class ProgressForm : Form
    {
        public bool UserCancelledMe { get; private set; }

        public ProgressForm(int PintMax, string PstrValue)
        {
            InitializeComponent();
            progressBar1.Maximum = PintMax;
            label1.Text = PstrValue;
            UserCancelledMe = false;
        }

        /// <summary>
        /// Sets the value of the label
        /// </summary>
        /// <param name="PstrMessage"></param>
        public void SetLabel(string PstrMessage)
        {
            label1.Text = PstrMessage;
        }

        /// <summary>
        /// Increments the progress bar
        /// </summary>
        public void Increment()
        {
            progressBar1.Increment(1);
        }

        /// <summary>
        /// sets the maximum value for the pb
        /// </summary>
        /// <param name="PintValue"></param>
        public void SetProgressBarMax(int PintValue)
        {
            progressBar1.Maximum = PintValue;
            progressBar1.Value = 0;
        }

        /// <summary>
        /// User clicked Cancel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult LobjResult = MessageBox.Show("Are you sure you want to cancel?",
                Common.APPNAME, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
            if (LobjResult == DialogResult.Yes)
            {
                UserCancelledMe = true;
            }
        }

        private void ProgressForm_Load(object sender, EventArgs e)
        {

        }
    }
}
