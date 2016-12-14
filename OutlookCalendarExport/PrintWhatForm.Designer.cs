namespace OutlookCalendarExport
{ 
    partial class PrintWhatForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PrintWhatForm));
            this.txtName = new System.Windows.Forms.TextBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.checkBoxHeader = new System.Windows.Forms.CheckBox();
            this.labelDate = new System.Windows.Forms.Label();
            this.dtPicker = new System.Windows.Forms.DateTimePicker();
            this.labelTemplate = new System.Windows.Forms.Label();
            this.listBoxTemplates = new System.Windows.Forms.ListBox();
            this.buttonDisplayOptions = new System.Windows.Forms.Button();
            this.groupBoxType = new System.Windows.Forms.GroupBox();
            this.radioButtonShared = new System.Windows.Forms.RadioButton();
            this.radioButtonMeetings = new System.Windows.Forms.RadioButton();
            this.radioButtonAll = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.checkBoxShowLocations = new System.Windows.Forms.CheckBox();
            this.checkBoxTimeOnly = new System.Windows.Forms.CheckBox();
            this.checkBoxRecurring = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxShow = new System.Windows.Forms.ComboBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.labelWhat = new System.Windows.Forms.Label();
            this.checkBoxExclude = new System.Windows.Forms.CheckBox();
            this.groupBoxType.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(147, 19);
            this.txtName.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtName.Name = "txtName";
            this.txtName.ReadOnly = true;
            this.txtName.Size = new System.Drawing.Size(874, 31);
            this.txtName.TabIndex = 1;
            this.txtName.TextChanged += new System.EventHandler(this.txtName_TextChanged);
            // 
            // btnOk
            // 
            this.btnOk.Enabled = false;
            this.btnOk.Location = new System.Drawing.Point(816, 773);
            this.btnOk.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(205, 44);
            this.btnOk.TabIndex = 15;
            this.btnOk.Text = "E&xport Now...";
            this.toolTip1.SetToolTip(this.btnOk, "Export now with the selected options.");
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(603, 773);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(205, 44);
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "Ca&ncel";
            this.toolTip1.SetToolTip(this.btnCancel, "Cancel and return ot Outlook.");
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // checkBoxHeader
            // 
            this.checkBoxHeader.AutoSize = true;
            this.checkBoxHeader.Checked = true;
            this.checkBoxHeader.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxHeader.Location = new System.Drawing.Point(651, 134);
            this.checkBoxHeader.Name = "checkBoxHeader";
            this.checkBoxHeader.Size = new System.Drawing.Size(279, 29);
            this.checkBoxHeader.TabIndex = 9;
            this.checkBoxHeader.Text = "Include names in &header";
            this.toolTip1.SetToolTip(this.checkBoxHeader, "Click to add the users names to the header of the page.\r\nIf you have set a displa" +
        "y name for the user, that will be used.\r\nIf you have not set a display name, the" +
        "ir full name will be used.\r\n");
            this.checkBoxHeader.UseVisualStyleBackColor = true;
            this.checkBoxHeader.CheckedChanged += new System.EventHandler(this.checkBoxHeader_CheckedChanged);
            // 
            // labelDate
            // 
            this.labelDate.AutoSize = true;
            this.labelDate.Location = new System.Drawing.Point(8, 66);
            this.labelDate.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelDate.Name = "labelDate";
            this.labelDate.Size = new System.Drawing.Size(63, 25);
            this.labelDate.TabIndex = 2;
            this.labelDate.Text = "Da&te:";
            // 
            // dtPicker
            // 
            this.dtPicker.Location = new System.Drawing.Point(147, 60);
            this.dtPicker.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dtPicker.Name = "dtPicker";
            this.dtPicker.Size = new System.Drawing.Size(490, 31);
            this.dtPicker.TabIndex = 3;
            this.toolTip1.SetToolTip(this.dtPicker, resources.GetString("dtPicker.ToolTip"));
            this.dtPicker.ValueChanged += new System.EventHandler(this.dtPicker_ValueChanged);
            // 
            // labelTemplate
            // 
            this.labelTemplate.AutoSize = true;
            this.labelTemplate.Location = new System.Drawing.Point(8, 310);
            this.labelTemplate.Name = "labelTemplate";
            this.labelTemplate.Size = new System.Drawing.Size(107, 25);
            this.labelTemplate.TabIndex = 13;
            this.labelTemplate.Text = "Te&mplate:";
            // 
            // listBoxTemplates
            // 
            this.listBoxTemplates.FormattingEnabled = true;
            this.listBoxTemplates.ItemHeight = 25;
            this.listBoxTemplates.Location = new System.Drawing.Point(147, 310);
            this.listBoxTemplates.Name = "listBoxTemplates";
            this.listBoxTemplates.Size = new System.Drawing.Size(490, 329);
            this.listBoxTemplates.TabIndex = 14;
            this.toolTip1.SetToolTip(this.listBoxTemplates, "Select a template to export to.");
            this.listBoxTemplates.SelectedIndexChanged += new System.EventHandler(this.listBoxTemplates_SelectedIndexChanged);
            // 
            // buttonDisplayOptions
            // 
            this.buttonDisplayOptions.Location = new System.Drawing.Point(13, 15);
            this.buttonDisplayOptions.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonDisplayOptions.Name = "buttonDisplayOptions";
            this.buttonDisplayOptions.Size = new System.Drawing.Size(126, 38);
            this.buttonDisplayOptions.TabIndex = 1;
            this.buttonDisplayOptions.Text = "&Names..";
            this.toolTip1.SetToolTip(this.buttonDisplayOptions, "Click to add, remove, or modify settings for users calendars you want exported.");
            this.buttonDisplayOptions.UseVisualStyleBackColor = true;
            this.buttonDisplayOptions.Click += new System.EventHandler(this.buttonDisplayOptions_Click);
            // 
            // groupBoxType
            // 
            this.groupBoxType.Controls.Add(this.radioButtonShared);
            this.groupBoxType.Controls.Add(this.radioButtonMeetings);
            this.groupBoxType.Controls.Add(this.radioButtonAll);
            this.groupBoxType.Location = new System.Drawing.Point(147, 99);
            this.groupBoxType.Name = "groupBoxType";
            this.groupBoxType.Size = new System.Drawing.Size(490, 153);
            this.groupBoxType.TabIndex = 15;
            this.groupBoxType.TabStop = false;
            this.groupBoxType.Text = "Export:";
            // 
            // radioButtonShared
            // 
            this.radioButtonShared.AutoSize = true;
            this.radioButtonShared.Location = new System.Drawing.Point(6, 101);
            this.radioButtonShared.Name = "radioButtonShared";
            this.radioButtonShared.Size = new System.Drawing.Size(252, 29);
            this.radioButtonShared.TabIndex = 6;
            this.radioButtonShared.TabStop = true;
            this.radioButtonShared.Text = "&Shared Meetings only";
            this.toolTip1.SetToolTip(this.radioButtonShared, "Select to export only meetings that one or more of the users are attending togeth" +
        "er");
            this.radioButtonShared.UseVisualStyleBackColor = true;
            this.radioButtonShared.CheckedChanged += new System.EventHandler(this.radioButtonShared_CheckedChanged);
            // 
            // radioButtonMeetings
            // 
            this.radioButtonMeetings.AutoSize = true;
            this.radioButtonMeetings.Location = new System.Drawing.Point(7, 66);
            this.radioButtonMeetings.Name = "radioButtonMeetings";
            this.radioButtonMeetings.Size = new System.Drawing.Size(177, 29);
            this.radioButtonMeetings.TabIndex = 5;
            this.radioButtonMeetings.TabStop = true;
            this.radioButtonMeetings.Text = "Meetings &only";
            this.toolTip1.SetToolTip(this.radioButtonMeetings, "Select to export only meetings from the users calendars");
            this.radioButtonMeetings.UseVisualStyleBackColor = true;
            this.radioButtonMeetings.CheckedChanged += new System.EventHandler(this.radioButtonMeetings_CheckedChanged);
            // 
            // radioButtonAll
            // 
            this.radioButtonAll.AutoSize = true;
            this.radioButtonAll.Checked = true;
            this.radioButtonAll.Location = new System.Drawing.Point(7, 30);
            this.radioButtonAll.Name = "radioButtonAll";
            this.radioButtonAll.Size = new System.Drawing.Size(213, 29);
            this.radioButtonAll.TabIndex = 4;
            this.radioButtonAll.TabStop = true;
            this.radioButtonAll.Text = "&All calendar items";
            this.toolTip1.SetToolTip(this.radioButtonAll, "Select this to export all appointments and meetings for all users calendars");
            this.radioButtonAll.UseVisualStyleBackColor = true;
            this.radioButtonAll.CheckedChanged += new System.EventHandler(this.radioButtonAll_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pictureBox1);
            this.groupBox1.Location = new System.Drawing.Point(651, 260);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(377, 379);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Preview:";
            // 
            // pictureBox1
            // 
            this.pictureBox1.AccessibleDescription = "No template selected to preview";
            this.pictureBox1.Location = new System.Drawing.Point(7, 31);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(353, 342);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // checkBoxShowLocations
            // 
            this.checkBoxShowLocations.AutoSize = true;
            this.checkBoxShowLocations.Checked = true;
            this.checkBoxShowLocations.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxShowLocations.Location = new System.Drawing.Point(651, 64);
            this.checkBoxShowLocations.Name = "checkBoxShowLocations";
            this.checkBoxShowLocations.Size = new System.Drawing.Size(287, 29);
            this.checkBoxShowLocations.TabIndex = 7;
            this.checkBoxShowLocations.Text = "&Include meeting locations";
            this.toolTip1.SetToolTip(this.checkBoxShowLocations, "Check to add meeting locations after the meeting subject.");
            this.checkBoxShowLocations.UseVisualStyleBackColor = true;
            this.checkBoxShowLocations.CheckedChanged += new System.EventHandler(this.checkBoxShowLocations_CheckedChanged);
            // 
            // checkBoxTimeOnly
            // 
            this.checkBoxTimeOnly.AutoSize = true;
            this.checkBoxTimeOnly.Location = new System.Drawing.Point(651, 99);
            this.checkBoxTimeOnly.Name = "checkBoxTimeOnly";
            this.checkBoxTimeOnly.Size = new System.Drawing.Size(218, 29);
            this.checkBoxTimeOnly.TabIndex = 8;
            this.checkBoxTimeOnly.Text = "&Display times only";
            this.toolTip1.SetToolTip(this.checkBoxTimeOnly, "Click to display only the time and replace the \r\nmeeting/appointment subjects wit" +
        "h a generic: \r\n -- Meeting -- or,\r\n -- Appointment --");
            this.checkBoxTimeOnly.UseVisualStyleBackColor = true;
            this.checkBoxTimeOnly.CheckedChanged += new System.EventHandler(this.checkBoxTimeOnly_CheckedChanged);
            // 
            // checkBoxRecurring
            // 
            this.checkBoxRecurring.AutoSize = true;
            this.checkBoxRecurring.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxRecurring.Location = new System.Drawing.Point(651, 169);
            this.checkBoxRecurring.Name = "checkBoxRecurring";
            this.checkBoxRecurring.Size = new System.Drawing.Size(298, 29);
            this.checkBoxRecurring.TabIndex = 10;
            this.checkBoxRecurring.Text = "&Emphasize recurring items";
            this.toolTip1.SetToolTip(this.checkBoxRecurring, "Click to add emphasis to recurring meetings so they \r\nare easier to identify once" +
        " exported.");
            this.checkBoxRecurring.UseVisualStyleBackColor = true;
            this.checkBoxRecurring.CheckedChanged += new System.EventHandler(this.checkBoxRecurring_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 260);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 25);
            this.label1.TabIndex = 11;
            this.label1.Text = "Sho&w:";
            // 
            // comboBoxShow
            // 
            this.comboBoxShow.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxShow.FormattingEnabled = true;
            this.comboBoxShow.Items.AddRange(new object[] {
            "All Templates",
            "Daily Templates",
            "Work week Templates",
            "Full week Templates",
            "Month Templates"});
            this.comboBoxShow.Location = new System.Drawing.Point(147, 260);
            this.comboBoxShow.Name = "comboBoxShow";
            this.comboBoxShow.Size = new System.Drawing.Size(490, 33);
            this.comboBoxShow.TabIndex = 12;
            this.toolTip1.SetToolTip(this.comboBoxShow, "Select to limit the types of templates displayed");
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(11, 773);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(205, 44);
            this.button1.TabIndex = 17;
            this.button1.Text = "Abo&ut...";
            this.toolTip1.SetToolTip(this.button1, "Show the about box for this add-in.");
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 651);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 25);
            this.label2.TabIndex = 18;
            this.label2.Text = "Exporting:";
            // 
            // labelWhat
            // 
            this.labelWhat.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.labelWhat.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelWhat.Location = new System.Drawing.Point(142, 651);
            this.labelWhat.Name = "labelWhat";
            this.labelWhat.Size = new System.Drawing.Size(869, 106);
            this.labelWhat.TabIndex = 19;
            this.labelWhat.Text = "Select options to see what will be exported";
            // 
            // checkBoxExclude
            // 
            this.checkBoxExclude.AutoSize = true;
            this.checkBoxExclude.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBoxExclude.Location = new System.Drawing.Point(651, 204);
            this.checkBoxExclude.Name = "checkBoxExclude";
            this.checkBoxExclude.Size = new System.Drawing.Size(327, 29);
            this.checkBoxExclude.TabIndex = 20;
            this.checkBoxExclude.Text = "Exclude &private appointments";
            this.toolTip1.SetToolTip(this.checkBoxExclude, "Click to add emphasis to recurring meetings so they \r\nare easier to identify once" +
        " exported.");
            this.checkBoxExclude.UseVisualStyleBackColor = true;
            // 
            // PrintWhatForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1044, 831);
            this.Controls.Add(this.checkBoxExclude);
            this.Controls.Add(this.labelWhat);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.comboBoxShow);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkBoxRecurring);
            this.Controls.Add(this.checkBoxTimeOnly);
            this.Controls.Add(this.checkBoxShowLocations);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBoxType);
            this.Controls.Add(this.listBoxTemplates);
            this.Controls.Add(this.labelTemplate);
            this.Controls.Add(this.checkBoxHeader);
            this.Controls.Add(this.buttonDisplayOptions);
            this.Controls.Add(this.labelDate);
            this.Controls.Add(this.dtPicker);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.txtName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PrintWhatForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Export Calendar to Word Options";
            this.Load += new System.EventHandler(this.PrintOptions_Load);
            this.groupBoxType.ResumeLayout(false);
            this.groupBoxType.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
    private System.Windows.Forms.CheckBox checkBoxHeader;
    private System.Windows.Forms.Label labelDate;
    private System.Windows.Forms.DateTimePicker dtPicker;
        private System.Windows.Forms.Label labelTemplate;
        private System.Windows.Forms.ListBox listBoxTemplates;
        private System.Windows.Forms.Button buttonDisplayOptions;
        private System.Windows.Forms.GroupBox groupBoxType;
        private System.Windows.Forms.RadioButton radioButtonShared;
        private System.Windows.Forms.RadioButton radioButtonMeetings;
        private System.Windows.Forms.RadioButton radioButtonAll;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox checkBoxShowLocations;
        private System.Windows.Forms.CheckBox checkBoxTimeOnly;
        internal System.Windows.Forms.CheckBox checkBoxRecurring;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxShow;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label labelWhat;
        internal System.Windows.Forms.CheckBox checkBoxExclude;
    }
}