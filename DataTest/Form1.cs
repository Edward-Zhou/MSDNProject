using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataTest
{
    public partial class Form1 : Form
    {
        protected SqlConnection sqlconn;
        protected SqlDataAdapter adapter;
        public Form1()
        {
            InitializeComponent();
            
        }

        private void CreateDataTable_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable("ParentTable");
            DataColumn column;
            DataRow row;
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "Id";
            column.ReadOnly = true;
            table.Columns.Add(column);
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Name";
            table.Columns.Add(column);
            DataSet dataset = new DataSet();
            dataset.Tables.Add(table);
            for (int i = 0; i <= 2; i++)
            {
                row = table.NewRow();
                row["Id"] = i;
                row["Name"] = "Name" + i;
                table.Rows.Add(row);
               // row.Delete();
               // MessageBox.Show(row.RowState.ToString());
            }
            table.AcceptChanges();
            row = table.Rows[0];
            row.Delete();
            
            MessageBox.Show(row.RowState.ToString());
        }

        DataSet dataset = new DataSet("User");
        private void ConnectData_Click(object sender, EventArgs e)
        {
            string sqlconnstr = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=D:\Sample\OfficeDev\AppsForOffice\AppsForOffice\DataTest\Data\DataDb.mdf;Integrated Security=True";

            sqlconn = new SqlConnection(sqlconnstr);

            try
            {

                sqlconn.Open();

                SqlCommand command = new SqlCommand("Select * From TestDb;", sqlconn);

                adapter = new SqlDataAdapter(command);

                SqlCommandBuilder cmdbuilder = new SqlCommandBuilder(adapter);

                command.CommandType = CommandType.Text;

                adapter.SelectCommand = command;

                adapter.Fill(dataset, "User");
                int RowCount = dataset.Tables["User"].Rows.Count;
                if (RowCount > 0)
                {
                    dataGridView1.DataSource = dataset.Tables["User"];
                }
                else
                {
                    dataGridView1.DataSource = null;
                }

                
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void UpdateTable_Click(object sender, EventArgs e)
        {
            DataTable dt = this.dataGridView1.DataSource as DataTable;

            if (dt != null)
            {
                //dt.AcceptChanges();
                adapter.Update(dt);
            }
        }

        private void CopyRow_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = this.dataGridView1.DataSource as DataTable;
            DataRow[] row= dt.AsEnumerable().Take(3).ToArray() ;
            DataTable dtnew = dt.Clone();
            dtnew.Clear();
            foreach (var item in row)
            {
                dtnew.Rows.Add(item.ItemArray);
            }
            this.dataGridView1.DataSource = dtnew;
        }

        private void CreateTable_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable("User");
            DataColumn dc = new DataColumn();
            dc.ColumnName = "Id";
            dc.DataType = Type.GetType("System.Int32");
            dt.Columns.Add(dc);
            dc = new DataColumn();
            dc.ColumnName = "Name";
            dc.DataType = Type.GetType("System.String");
            dt.Columns.Add(dc);
            for (int i = 0; i < 5; i++)
            {
                dt.Rows.Add(i, "Name" + i.ToString());
            }
            this.dataGridView1.DataSource = dt;
            this.dataGridView1.TopLeftHeaderCell.Value = "Hello";
            
        }

        private void DeleteRowValue_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = this.dataGridView1.DataSource as DataTable;
            DataRow dr = dt.Rows[0];
            dr.Delete();
            if (dr.RowState == DataRowState.Deleted)
            {
                string name = dr["Name", DataRowVersion.Original].ToString();
                MessageBox.Show(name);
            }
            //delete row

            
        }

        private void RowState_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataColumn dc = new DataColumn("Name", Type.GetType("System.String"));
            dt.Columns.Add(dc);
            DataRow dr ;
            dr = dt.NewRow();
            MessageBox.Show("New Row " + dr.RowState.ToString());
            dt.Rows.Add(dr);
            MessageBox.Show("AddRow "+ dr.RowState.ToString());
            dt.AcceptChanges();
            MessageBox.Show("AcceptChanges "+dr.RowState.ToString());
            dr["Name"] = "Name";
            MessageBox.Show("Modified " + dr.RowState.ToString()); 
            dr.Delete();
            MessageBox.Show("Deleted " + dr.RowState);
            if (dr.RowState == DataRowState.Deleted)
            {
                MessageBox.Show("Deleted");
            }
        }
    }
}
