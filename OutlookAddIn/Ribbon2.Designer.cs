namespace OutlookAddIn
{
    partial class Ribbon2 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon2()
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
            this.MailStatus = this.Factory.CreateRibbonButton();
            this.WordRange = this.Factory.CreateRibbonButton();
            this.ItemAddbtn = this.Factory.CreateRibbonButton();
            this.ShowForm = this.Factory.CreateRibbonButton();
            this.ShowWebView = this.Factory.CreateRibbonButton();
            this.SetPropertybtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.MailStatus);
            this.group1.Items.Add(this.WordRange);
            this.group1.Items.Add(this.ItemAddbtn);
            this.group1.Items.Add(this.ShowForm);
            this.group1.Items.Add(this.ShowWebView);
            this.group1.Items.Add(this.SetPropertybtn);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // MailStatus
            // 
            this.MailStatus.Label = "MailStatus";
            this.MailStatus.Name = "MailStatus";
            this.MailStatus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MailStatus_Click);
            // 
            // WordRange
            // 
            this.WordRange.Label = "WordRange";
            this.WordRange.Name = "WordRange";
            this.WordRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WordRange_Click);
            // 
            // ItemAddbtn
            // 
            this.ItemAddbtn.Label = "ItemAddbtn";
            this.ItemAddbtn.Name = "ItemAddbtn";
            this.ItemAddbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ItemAddbtn_Click);
            // 
            // ShowForm
            // 
            this.ShowForm.Label = "ShowForm";
            this.ShowForm.Name = "ShowForm";
            this.ShowForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowForm_Click);
            // 
            // ShowWebView
            // 
            this.ShowWebView.Label = "ShowWebView";
            this.ShowWebView.Name = "ShowWebView";
            this.ShowWebView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowWebView_Click);
            // 
            // SetPropertybtn
            // 
            this.SetPropertybtn.Label = "SetProperty";
            this.SetPropertybtn.Name = "SetPropertybtn";
            this.SetPropertybtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetPropertybtn_Click);
            // 
            // Ribbon2
            // 
            this.Name = "Ribbon2";
            this.RibbonType = "Microsoft.Outlook.Appointment, Microsoft.Outlook.Contact, Microsoft.Outlook.Explo" +
    "rer, Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon2_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MailStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WordRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ItemAddbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowWebView;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetPropertybtn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon2 Ribbon2
        {
            get { return this.GetRibbon<Ribbon2>(); }
        }
    }
}
