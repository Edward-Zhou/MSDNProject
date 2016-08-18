namespace DataTest
{
    partial class Form1
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
            this.CreateDataTable = new System.Windows.Forms.Button();
            this.ConnectData = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.UpdateTable = new System.Windows.Forms.Button();
            this.CopyRow = new System.Windows.Forms.Button();
            this.CreateTable = new System.Windows.Forms.Button();
            this.DeleteRowValue = new System.Windows.Forms.Button();
            this.RowState = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // CreateDataTable
            // 
            this.CreateDataTable.Location = new System.Drawing.Point(12, 12);
            this.CreateDataTable.Name = "CreateDataTable";
            this.CreateDataTable.Size = new System.Drawing.Size(101, 23);
            this.CreateDataTable.TabIndex = 0;
            this.CreateDataTable.Text = "CreateDataTable";
            this.CreateDataTable.UseVisualStyleBackColor = true;
            this.CreateDataTable.Click += new System.EventHandler(this.CreateDataTable_Click);
            // 
            // ConnectData
            // 
            this.ConnectData.Location = new System.Drawing.Point(119, 12);
            this.ConnectData.Name = "ConnectData";
            this.ConnectData.Size = new System.Drawing.Size(101, 23);
            this.ConnectData.TabIndex = 1;
            this.ConnectData.Text = "ConnectData";
            this.ConnectData.UseVisualStyleBackColor = true;
            this.ConnectData.Click += new System.EventHandler(this.ConnectData_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(0, 122);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(357, 150);
            this.dataGridView1.TabIndex = 2;
            // 
            // UpdateTable
            // 
            this.UpdateTable.Location = new System.Drawing.Point(226, 12);
            this.UpdateTable.Name = "UpdateTable";
            this.UpdateTable.Size = new System.Drawing.Size(101, 23);
            this.UpdateTable.TabIndex = 3;
            this.UpdateTable.Text = "UpdateTable";
            this.UpdateTable.UseVisualStyleBackColor = true;
            this.UpdateTable.Click += new System.EventHandler(this.UpdateTable_Click);
            // 
            // CopyRow
            // 
            this.CopyRow.Location = new System.Drawing.Point(119, 41);
            this.CopyRow.Name = "CopyRow";
            this.CopyRow.Size = new System.Drawing.Size(101, 23);
            this.CopyRow.TabIndex = 4;
            this.CopyRow.Text = "CopyRow";
            this.CopyRow.UseVisualStyleBackColor = true;
            this.CopyRow.Click += new System.EventHandler(this.CopyRow_Click);
            // 
            // CreateTable
            // 
            this.CreateTable.Location = new System.Drawing.Point(12, 41);
            this.CreateTable.Name = "CreateTable";
            this.CreateTable.Size = new System.Drawing.Size(101, 23);
            this.CreateTable.TabIndex = 5;
            this.CreateTable.Text = "CreateTable";
            this.CreateTable.UseVisualStyleBackColor = true;
            this.CreateTable.Click += new System.EventHandler(this.CreateTable_Click);
            // 
            // DeleteRowValue
            // 
            this.DeleteRowValue.Location = new System.Drawing.Point(226, 41);
            this.DeleteRowValue.Name = "DeleteRowValue";
            this.DeleteRowValue.Size = new System.Drawing.Size(101, 23);
            this.DeleteRowValue.TabIndex = 6;
            this.DeleteRowValue.Text = "DeleteRowValue";
            this.DeleteRowValue.UseVisualStyleBackColor = true;
            this.DeleteRowValue.Click += new System.EventHandler(this.DeleteRowValue_Click);
            // 
            // RowState
            // 
            this.RowState.Location = new System.Drawing.Point(12, 70);
            this.RowState.Name = "RowState";
            this.RowState.Size = new System.Drawing.Size(101, 23);
            this.RowState.TabIndex = 7;
            this.RowState.Text = "RowState";
            this.RowState.UseVisualStyleBackColor = true;
            this.RowState.Click += new System.EventHandler(this.RowState_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(359, 272);
            this.Controls.Add(this.RowState);
            this.Controls.Add(this.DeleteRowValue);
            this.Controls.Add(this.CreateTable);
            this.Controls.Add(this.CopyRow);
            this.Controls.Add(this.UpdateTable);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.ConnectData);
            this.Controls.Add(this.CreateDataTable);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button CreateDataTable;
        private System.Windows.Forms.Button ConnectData;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button UpdateTable;
        private System.Windows.Forms.Button CopyRow;
        private System.Windows.Forms.Button CreateTable;
        private System.Windows.Forms.Button DeleteRowValue;
        private System.Windows.Forms.Button RowState;
    }
}

