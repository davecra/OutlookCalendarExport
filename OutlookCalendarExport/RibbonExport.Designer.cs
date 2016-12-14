namespace OutlookCalendarExport
{
    partial class RibbonExport : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonExport()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonExport = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabCalendar";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabCalendar";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonExport);
            this.group1.Label = "Export";
            this.group1.Name = "group1";
            // 
            // buttonExport
            // 
            this.buttonExport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonExport.Image = global::OutlookCalendarExport.Properties.Resources.day;
            this.buttonExport.Label = "Export Calendar to Word";
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.OfficeImageId = "ControlLayoutStacked";
            this.buttonExport.ShowImage = true;
            this.buttonExport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonExport_Click);
            // 
            // RibbonExport
            // 
            this.Name = "RibbonExport";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonExport;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonExport Ribbon1
        {
            get { return this.GetRibbon<RibbonExport>(); }
        }
    }
}
