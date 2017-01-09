namespace PowerPointAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.NewSlide = this.Factory.CreateRibbonButton();
            this.ExportSlide = this.Factory.CreateRibbonButton();
            this.addTaskPane = this.Factory.CreateRibbonButton();
            this.txtRange = this.Factory.CreateRibbonButton();
            this.TaskPaneWindows = this.Factory.CreateRibbonButton();
            this.InsertImg = this.Factory.CreateRibbonButton();
            this.Quit = this.Factory.CreateRibbonButton();
            this.Showbtn = this.Factory.CreateRibbonButton();
            this.ShowDialogbtn = this.Factory.CreateRibbonButton();
            this.GetSelection = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.NewSlide);
            this.group1.Items.Add(this.ExportSlide);
            this.group1.Items.Add(this.addTaskPane);
            this.group1.Items.Add(this.txtRange);
            this.group1.Items.Add(this.TaskPaneWindows);
            this.group1.Items.Add(this.InsertImg);
            this.group1.Items.Add(this.Quit);
            this.group1.Items.Add(this.Showbtn);
            this.group1.Items.Add(this.ShowDialogbtn);
            this.group1.Items.Add(this.GetSelection);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // NewSlide
            // 
            this.NewSlide.Label = "NewSlide";
            this.NewSlide.Name = "NewSlide";
            this.NewSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.NewSlide_Click);
            // 
            // ExportSlide
            // 
            this.ExportSlide.Label = "ExportSlide";
            this.ExportSlide.Name = "ExportSlide";
            this.ExportSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportSlide_Click);
            // 
            // addTaskPane
            // 
            this.addTaskPane.Label = "addTaskPane";
            this.addTaskPane.Name = "addTaskPane";
            this.addTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addTaskPane_Click);
            // 
            // txtRange
            // 
            this.txtRange.Label = "txtRange";
            this.txtRange.Name = "txtRange";
            this.txtRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.txtRange_Click);
            // 
            // TaskPaneWindows
            // 
            this.TaskPaneWindows.Label = "TaskPaneWindows";
            this.TaskPaneWindows.Name = "TaskPaneWindows";
            this.TaskPaneWindows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TaskPaneWindows_Click);
            // 
            // InsertImg
            // 
            this.InsertImg.Label = "InsertImg";
            this.InsertImg.Name = "InsertImg";
            this.InsertImg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertImg_Click);
            // 
            // Quit
            // 
            this.Quit.Label = "Quit";
            this.Quit.Name = "Quit";
            this.Quit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Quit_Click);
            // 
            // Showbtn
            // 
            this.Showbtn.Label = "Show";
            this.Showbtn.Name = "Showbtn";
            this.Showbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Showbtn_Click);
            // 
            // ShowDialogbtn
            // 
            this.ShowDialogbtn.Label = "ShowDialog";
            this.ShowDialogbtn.Name = "ShowDialogbtn";
            this.ShowDialogbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowDialogbtn_Click);
            // 
            // GetSelection
            // 
            this.GetSelection.Label = "GetSelection";
            this.GetSelection.Name = "GetSelection";
            this.GetSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetSelection_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton NewSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton txtRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TaskPaneWindows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertImg;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Quit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Showbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowDialogbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetSelection;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
