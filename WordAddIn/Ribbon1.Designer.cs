﻿namespace WordAddIn
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
            this.ShapeFormat = this.Factory.CreateRibbonButton();
            this.ParaIndentbtn = this.Factory.CreateRibbonButton();
            this.addContentControl = this.Factory.CreateRibbonButton();
            this.WrapTable = this.Factory.CreateRibbonButton();
            this.SelectionMove = this.Factory.CreateRibbonButton();
            this.LogBtn = this.Factory.CreateRibbonButton();
            this.ChangeLogPath = this.Factory.CreateRibbonButton();
            this.changeLocation = this.Factory.CreateRibbonButton();
            this.findReplace = this.Factory.CreateRibbonButton();
            this.InsertContentControl = this.Factory.CreateRibbonButton();
            this.RangeReplace = this.Factory.CreateRibbonButton();
            this.TableCell = this.Factory.CreateRibbonButton();
            this.AddInName = this.Factory.CreateRibbonButton();
            this.SaveAsTemplate = this.Factory.CreateRibbonButton();
            this.headerFooter = this.Factory.CreateRibbonButton();
            this.MoveShape = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.ShapeFormat);
            this.group1.Items.Add(this.ParaIndentbtn);
            this.group1.Items.Add(this.addContentControl);
            this.group1.Items.Add(this.WrapTable);
            this.group1.Items.Add(this.SelectionMove);
            this.group1.Items.Add(this.LogBtn);
            this.group1.Items.Add(this.ChangeLogPath);
            this.group1.Items.Add(this.changeLocation);
            this.group1.Items.Add(this.findReplace);
            this.group1.Items.Add(this.InsertContentControl);
            this.group1.Items.Add(this.RangeReplace);
            this.group1.Items.Add(this.TableCell);
            this.group1.Items.Add(this.AddInName);
            this.group1.Items.Add(this.SaveAsTemplate);
            this.group1.Items.Add(this.headerFooter);
            this.group1.Items.Add(this.MoveShape);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // ShapeFormat
            // 
            this.ShapeFormat.Label = "ShapeFormat";
            this.ShapeFormat.Name = "ShapeFormat";
            this.ShapeFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShapeFormat_Click);
            // 
            // ParaIndentbtn
            // 
            this.ParaIndentbtn.Label = "ParaIndent";
            this.ParaIndentbtn.Name = "ParaIndentbtn";
            this.ParaIndentbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ParaIndentbtn_Click);
            // 
            // addContentControl
            // 
            this.addContentControl.Label = "addContentControl";
            this.addContentControl.Name = "addContentControl";
            this.addContentControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addContentControl_Click);
            // 
            // WrapTable
            // 
            this.WrapTable.Label = "WrapTable";
            this.WrapTable.Name = "WrapTable";
            this.WrapTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WrapTable_Click);
            // 
            // SelectionMove
            // 
            this.SelectionMove.Label = "SelectionMove";
            this.SelectionMove.Name = "SelectionMove";
            this.SelectionMove.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectionMove_Click);
            // 
            // LogBtn
            // 
            this.LogBtn.Label = "Log";
            this.LogBtn.Name = "LogBtn";
            this.LogBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.LogBtn_Click);
            // 
            // ChangeLogPath
            // 
            this.ChangeLogPath.Label = "ChangeLogPath";
            this.ChangeLogPath.Name = "ChangeLogPath";
            this.ChangeLogPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ChangeLogPath_Click);
            // 
            // changeLocation
            // 
            this.changeLocation.Label = "changeLocation";
            this.changeLocation.Name = "changeLocation";
            this.changeLocation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changeLocation_Click);
            // 
            // findReplace
            // 
            this.findReplace.Label = "findReplace";
            this.findReplace.Name = "findReplace";
            this.findReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.findReplace_Click);
            // 
            // InsertContentControl
            // 
            this.InsertContentControl.Label = "InsertContentControl";
            this.InsertContentControl.Name = "InsertContentControl";
            this.InsertContentControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.InsertContentControl_Click);
            // 
            // RangeReplace
            // 
            this.RangeReplace.Label = "RangeReplace";
            this.RangeReplace.Name = "RangeReplace";
            this.RangeReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RangeReplace_Click);
            // 
            // TableCell
            // 
            this.TableCell.Label = "TableCell";
            this.TableCell.Name = "TableCell";
            this.TableCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TableCell_Click);
            // 
            // AddInName
            // 
            this.AddInName.Label = "AddInName";
            this.AddInName.Name = "AddInName";
            this.AddInName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddInName_Click);
            // 
            // SaveAsTemplate
            // 
            this.SaveAsTemplate.Label = "SaveAsTemplate";
            this.SaveAsTemplate.Name = "SaveAsTemplate";
            this.SaveAsTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SaveAsTemplate_Click);
            // 
            // headerFooter
            // 
            this.headerFooter.Label = "headerFooter";
            this.headerFooter.Name = "headerFooter";
            this.headerFooter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.headerFooter_Click);
            // 
            // MoveShape
            // 
            this.MoveShape.Label = "MoveShape";
            this.MoveShape.Name = "MoveShape";
            this.MoveShape.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MoveShape_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ShapeFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ParaIndentbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addContentControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton WrapTable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SelectionMove;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton LogBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ChangeLogPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton changeLocation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton findReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton InsertContentControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RangeReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TableCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddInName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SaveAsTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton headerFooter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MoveShape;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}