namespace ExcelAddIn
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.CreateShape = this.Factory.CreateRibbonButton();
            this.ExcelTemplate = this.Factory.CreateRibbonButton();
            this.Classbtn = this.Factory.CreateRibbonButton();
            this.CopyPivotTable = this.Factory.CreateRibbonButton();
            this.CallMacro = this.Factory.CreateRibbonButton();
            this.qryTables = this.Factory.CreateRibbonButton();
            this.WorkBookSize = this.Factory.CreateRibbonButton();
            this.SelectUsedRange = this.Factory.CreateRibbonButton();
            this.ExcelCopybtn = this.Factory.CreateRibbonButton();
            this.ShowTaskPane = this.Factory.CreateRibbonButton();
            this.AllowEditRangesbtn = this.Factory.CreateRibbonButton();
            this.DeleteRange = this.Factory.CreateRibbonButton();
            this.FilePre = this.Factory.CreateRibbonButton();
            this.ShapeSave = this.Factory.CreateRibbonButton();
            this.CustomTaskPane = this.Factory.CreateRibbonButton();
            this.AddTaskPane = this.Factory.CreateRibbonButton();
            this.AddSecondTaskPane = this.Factory.CreateRibbonButton();
            this.TaskCount = this.Factory.CreateRibbonButton();
            this.AddDataLabel = this.Factory.CreateRibbonButton();
            this.ChangeDataLabel = this.Factory.CreateRibbonButton();
            this.DataLabelPosition = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.UnSaveBtn = this.Factory.CreateRibbonButton();
            this.PivotFilter = this.Factory.CreateRibbonButton();
            this.TopFilter = this.Factory.CreateRibbonButton();
            this.showForm = this.Factory.CreateRibbonButton();
            this.comboBox1 = this.Factory.CreateRibbonComboBox();
            this.RegisterEvent = this.Factory.CreateRibbonButton();
            this.ExportToPdf = this.Factory.CreateRibbonButton();
            this.AddHyperlink = this.Factory.CreateRibbonButton();
            this.CreateList = this.Factory.CreateRibbonButton();
            this.ActiveTab = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.CreateShape);
            this.group1.Items.Add(this.ExcelTemplate);
            this.group1.Items.Add(this.Classbtn);
            this.group1.Items.Add(this.CopyPivotTable);
            this.group1.Items.Add(this.CallMacro);
            this.group1.Items.Add(this.qryTables);
            this.group1.Items.Add(this.WorkBookSize);
            this.group1.Items.Add(this.SelectUsedRange);
            this.group1.Items.Add(this.ExcelCopybtn);
            this.group1.Items.Add(this.ShowTaskPane);
            this.group1.Items.Add(this.AllowEditRangesbtn);
            this.group1.Items.Add(this.DeleteRange);
            this.group1.Items.Add(this.FilePre);
            this.group1.Items.Add(this.ShapeSave);
            this.group1.Items.Add(this.CustomTaskPane);
            this.group1.Items.Add(this.AddTaskPane);
            this.group1.Items.Add(this.AddSecondTaskPane);
            this.group1.Items.Add(this.TaskCount);
            this.group1.Items.Add(this.AddDataLabel);
            this.group1.Items.Add(this.ChangeDataLabel);
            this.group1.Items.Add(this.DataLabelPosition);
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.UnSaveBtn);
            this.group1.Items.Add(this.PivotFilter);
            this.group1.Items.Add(this.TopFilter);
            this.group1.Items.Add(this.showForm);
            this.group1.Items.Add(this.comboBox1);
            this.group1.Items.Add(this.RegisterEvent);
            this.group1.Items.Add(this.ExportToPdf);
            this.group1.Items.Add(this.AddHyperlink);
            this.group1.Items.Add(this.CreateList);
            this.group1.Items.Add(this.ActiveTab);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // CreateShape
            // 
            this.CreateShape.Label = "CreateShape";
            this.CreateShape.Name = "CreateShape";
            this.CreateShape.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateShape_Click);
            // 
            // ExcelTemplate
            // 
            this.ExcelTemplate.Label = "ExcelTemplate";
            this.ExcelTemplate.Name = "ExcelTemplate";
            this.ExcelTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExcelTemplate_Click);
            // 
            // Classbtn
            // 
            this.Classbtn.Label = "Classbtn";
            this.Classbtn.Name = "Classbtn";
            this.Classbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Classbtn_Click);
            // 
            // CopyPivotTable
            // 
            this.CopyPivotTable.Label = "CopyPivotTable";
            this.CopyPivotTable.Name = "CopyPivotTable";
            this.CopyPivotTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CopyPivotTable_Click);
            // 
            // CallMacro
            // 
            this.CallMacro.Label = "CallMacro";
            this.CallMacro.Name = "CallMacro";
            this.CallMacro.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CallMacro_Click);
            // 
            // qryTables
            // 
            this.qryTables.Label = "qryTables";
            this.qryTables.Name = "qryTables";
            this.qryTables.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.qryTables_Click);
            // 
            // WorkBookSize
            // 
            this.WorkBookSize.Label = "WorkBookSize";
            this.WorkBookSize.Name = "WorkBookSize";
            this.WorkBookSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WorkBookSize_Click);
            // 
            // SelectUsedRange
            // 
            this.SelectUsedRange.Label = "SelectUsedRange";
            this.SelectUsedRange.Name = "SelectUsedRange";
            this.SelectUsedRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectUsedRange_Click);
            // 
            // ExcelCopybtn
            // 
            this.ExcelCopybtn.Label = "ExcelCopy";
            this.ExcelCopybtn.Name = "ExcelCopybtn";
            this.ExcelCopybtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExcelCopybtn_Click);
            // 
            // ShowTaskPane
            // 
            this.ShowTaskPane.Label = "ShowTaskPane";
            this.ShowTaskPane.Name = "ShowTaskPane";
            this.ShowTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowTaskPane_Click);
            // 
            // AllowEditRangesbtn
            // 
            this.AllowEditRangesbtn.Label = "AllowEditRanges";
            this.AllowEditRangesbtn.Name = "AllowEditRangesbtn";
            this.AllowEditRangesbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AllowEditRangesbtn_Click);
            // 
            // DeleteRange
            // 
            this.DeleteRange.Label = "DeleteRange";
            this.DeleteRange.Name = "DeleteRange";
            this.DeleteRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DeleteRange_Click);
            // 
            // FilePre
            // 
            this.FilePre.Label = "FilePre";
            this.FilePre.Name = "FilePre";
            this.FilePre.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FilePre_Click);
            // 
            // ShapeSave
            // 
            this.ShapeSave.Label = "ShapeSave";
            this.ShapeSave.Name = "ShapeSave";
            this.ShapeSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShapeSave_Click);
            // 
            // CustomTaskPane
            // 
            this.CustomTaskPane.Label = "CustomTaskPane";
            this.CustomTaskPane.Name = "CustomTaskPane";
            this.CustomTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CustomTaskPane_Click);
            // 
            // AddTaskPane
            // 
            this.AddTaskPane.Label = "AddTaskPane";
            this.AddTaskPane.Name = "AddTaskPane";
            this.AddTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddTaskPane_Click);
            // 
            // AddSecondTaskPane
            // 
            this.AddSecondTaskPane.Label = "AddSecondTaskPane";
            this.AddSecondTaskPane.Name = "AddSecondTaskPane";
            this.AddSecondTaskPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddSecondTaskPane_Click);
            // 
            // TaskCount
            // 
            this.TaskCount.Label = "TaskCount";
            this.TaskCount.Name = "TaskCount";
            this.TaskCount.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TaskCount_Click);
            // 
            // AddDataLabel
            // 
            this.AddDataLabel.Label = "AddDataLabel";
            this.AddDataLabel.Name = "AddDataLabel";
            this.AddDataLabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddDataLabel_Click);
            // 
            // ChangeDataLabel
            // 
            this.ChangeDataLabel.Label = "ChangeDataLabel";
            this.ChangeDataLabel.Name = "ChangeDataLabel";
            this.ChangeDataLabel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeDataLabel_Click);
            // 
            // DataLabelPosition
            // 
            this.DataLabelPosition.Label = "DataLabelPosition";
            this.DataLabelPosition.Name = "DataLabelPosition";
            this.DataLabelPosition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DataLabelPosition_Click);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "分割";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // UnSaveBtn
            // 
            this.UnSaveBtn.Label = "UnSave";
            this.UnSaveBtn.Name = "UnSaveBtn";
            this.UnSaveBtn.ScreenTip = "Test";
            this.UnSaveBtn.SuperTip = "Test1";
            this.UnSaveBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UnSaveBtn_Click);
            // 
            // PivotFilter
            // 
            this.PivotFilter.Label = "PivotFilter";
            this.PivotFilter.Name = "PivotFilter";
            this.PivotFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.PivotFilter_Click);
            // 
            // TopFilter
            // 
            this.TopFilter.Label = "TopFilter";
            this.TopFilter.Name = "TopFilter";
            this.TopFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TopFilter_Click);
            // 
            // showForm
            // 
            this.showForm.Label = "showForm";
            this.showForm.Name = "showForm";
            this.showForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showForm_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.Label = "comboBox1";
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.ShowLabel = false;
            this.comboBox1.Text = null;
            // 
            // RegisterEvent
            // 
            this.RegisterEvent.Label = "RegisterEvent";
            this.RegisterEvent.Name = "RegisterEvent";
            this.RegisterEvent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RegisterEvent_Click);
            // 
            // ExportToPdf
            // 
            this.ExportToPdf.Label = "ExportToPdf";
            this.ExportToPdf.Name = "ExportToPdf";
            this.ExportToPdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportToPdf_Click);
            // 
            // AddHyperlink
            // 
            this.AddHyperlink.Label = "AddHyperlink";
            this.AddHyperlink.Name = "AddHyperlink";
            this.AddHyperlink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddHyperlink_Click);
            // 
            // CreateList
            // 
            this.CreateList.Label = "CreateList";
            this.CreateList.Name = "CreateList";
            this.CreateList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateList_Click);
            // 
            // ActiveTab
            // 
            this.ActiveTab.Label = "ActiveTab";
            this.ActiveTab.Name = "ActiveTab";
            this.ActiveTab.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ActiveTab_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateShape;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExcelTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Classbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CopyPivotTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CallMacro;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton qryTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WorkBookSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SelectUsedRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExcelCopybtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShowTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AllowEditRangesbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DeleteRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FilePre;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShapeSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CustomTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddSecondTaskPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TaskCount;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeDataLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddDataLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DataLabelPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UnSaveBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PivotFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TopFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RegisterEvent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportToPdf;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddHyperlink;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ActiveTab;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
