namespace WinForm.Forms
{
    partial class DataGridViewfrm
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.LoadCheckBoxHeader = new System.Windows.Forms.Button();
            this.AddComboxColumn = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.mSDNDataSet = new WinForm.MSDNDataSet();
            this.mSDNDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tableBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tableTableAdapter = new WinForm.MSDNDataSetTableAdapters.TableTableAdapter();
            this.idDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.uNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.comBobox1DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.comBobox2DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mSDNDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mSDNDataSetBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tableBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 37);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(240, 150);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_ColumnHeaderMouseClick);
            this.dataGridView1.Click += new System.EventHandler(this.dataGridView1_Click);
            // 
            // LoadCheckBoxHeader
            // 
            this.LoadCheckBoxHeader.Location = new System.Drawing.Point(12, 8);
            this.LoadCheckBoxHeader.Name = "LoadCheckBoxHeader";
            this.LoadCheckBoxHeader.Size = new System.Drawing.Size(127, 23);
            this.LoadCheckBoxHeader.TabIndex = 1;
            this.LoadCheckBoxHeader.Text = "LoadCheckBoxHeader";
            this.LoadCheckBoxHeader.UseVisualStyleBackColor = true;
            this.LoadCheckBoxHeader.Click += new System.EventHandler(this.LoadCheckBoxHeader_Click);
            // 
            // AddComboxColumn
            // 
            this.AddComboxColumn.Location = new System.Drawing.Point(145, 8);
            this.AddComboxColumn.Name = "AddComboxColumn";
            this.AddComboxColumn.Size = new System.Drawing.Size(127, 23);
            this.AddComboxColumn.TabIndex = 2;
            this.AddComboxColumn.Text = "AddComboxColumn";
            this.AddComboxColumn.UseVisualStyleBackColor = true;
            this.AddComboxColumn.Click += new System.EventHandler(this.AddComboxColumn_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AutoGenerateColumns = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.idDataGridViewTextBoxColumn,
            this.uNameDataGridViewTextBoxColumn,
            this.comBobox1DataGridViewTextBoxColumn,
            this.comBobox2DataGridViewTextBoxColumn});
            this.dataGridView2.DataSource = this.tableBindingSource;
            this.dataGridView2.Location = new System.Drawing.Point(331, 37);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(457, 150);
            this.dataGridView2.TabIndex = 3;
            // 
            // mSDNDataSet
            // 
            this.mSDNDataSet.DataSetName = "MSDNDataSet";
            this.mSDNDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // mSDNDataSetBindingSource
            // 
            this.mSDNDataSetBindingSource.DataSource = this.mSDNDataSet;
            this.mSDNDataSetBindingSource.Position = 0;
            // 
            // tableBindingSource
            // 
            this.tableBindingSource.DataMember = "Table";
            this.tableBindingSource.DataSource = this.mSDNDataSetBindingSource;
            // 
            // tableTableAdapter
            // 
            this.tableTableAdapter.ClearBeforeFill = true;
            // 
            // idDataGridViewTextBoxColumn
            // 
            this.idDataGridViewTextBoxColumn.DataPropertyName = "Id";
            this.idDataGridViewTextBoxColumn.HeaderText = "Id";
            this.idDataGridViewTextBoxColumn.Name = "idDataGridViewTextBoxColumn";
            // 
            // uNameDataGridViewTextBoxColumn
            // 
            this.uNameDataGridViewTextBoxColumn.DataPropertyName = "UName";
            this.uNameDataGridViewTextBoxColumn.HeaderText = "UName";
            this.uNameDataGridViewTextBoxColumn.Name = "uNameDataGridViewTextBoxColumn";
            // 
            // comBobox1DataGridViewTextBoxColumn
            // 
            this.comBobox1DataGridViewTextBoxColumn.DataPropertyName = "ComBobox1";
            this.comBobox1DataGridViewTextBoxColumn.DataSource = this.mSDNDataSetBindingSource;
            this.comBobox1DataGridViewTextBoxColumn.HeaderText = "ComBobox1";
            this.comBobox1DataGridViewTextBoxColumn.Name = "comBobox1DataGridViewTextBoxColumn";
            this.comBobox1DataGridViewTextBoxColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.comBobox1DataGridViewTextBoxColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // comBobox2DataGridViewTextBoxColumn
            // 
            this.comBobox2DataGridViewTextBoxColumn.DataPropertyName = "ComBobox2";
            this.comBobox2DataGridViewTextBoxColumn.HeaderText = "ComBobox2";
            this.comBobox2DataGridViewTextBoxColumn.Name = "comBobox2DataGridViewTextBoxColumn";
            // 
            // DataGridViewfrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(827, 384);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.AddComboxColumn);
            this.Controls.Add(this.LoadCheckBoxHeader);
            this.Controls.Add(this.dataGridView1);
            this.Name = "DataGridViewfrm";
            this.Text = "DataGridViewfrm";
            this.Load += new System.EventHandler(this.DataGridViewfrm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mSDNDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mSDNDataSetBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tableBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button LoadCheckBoxHeader;
        private System.Windows.Forms.Button AddComboxColumn;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.BindingSource mSDNDataSetBindingSource;
        private MSDNDataSet mSDNDataSet;
        private System.Windows.Forms.BindingSource tableBindingSource;
        private MSDNDataSetTableAdapters.TableTableAdapter tableTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn idDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn uNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewComboBoxColumn comBobox1DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn comBobox2DataGridViewTextBoxColumn;
    }
}