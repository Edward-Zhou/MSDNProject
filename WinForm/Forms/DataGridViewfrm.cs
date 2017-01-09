using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinForm.ClassLibrary;

namespace WinForm.Forms
{
    public partial class DataGridViewfrm : Form
    {
        public DataGridViewfrm()
        {
            InitializeComponent();
        }

        private void LoadCheckBoxHeader_Click(object sender, EventArgs e)
        {
            DataGridViewCheckBoxColumn colCB = new DataGridViewCheckBoxColumn();
            DatagridViewCheckBoxHeaderCell cbHeader = new DatagridViewCheckBoxHeaderCell();
            colCB.HeaderCell = cbHeader;
            dataGridView1.Columns.Add(colCB);
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //e.
        }
        int RowSelectCount=0;
        private void dataGridView1_Click(object sender, EventArgs e)
        {            
            
            if (RowSelectCount > 0)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.Selected = false;
                }
            }
            RowSelectCount = dataGridView1.SelectedRows.Count;
        }

        private void AddComboxColumn_Click(object sender, EventArgs e)
        {
            List<Produto> listProdutos = new List<Produto>();
            listProdutos.Add(new Produto() { Id = 1, Nome = "Produto 1" });
            listProdutos.Add(new Produto() { Id = 2, Nome = "Produto 2" });
            listProdutos.Add(new Produto() { Id = 3, Nome = "Produto 3" });
            listProdutos.Add(new Produto() { Id = 4, Nome = "Produto 4" });

            DataGridViewComboBoxColumn comboBoxColumn = new DataGridViewComboBoxColumn();
            comboBoxColumn.DataSource = listProdutos.ToList();
            comboBoxColumn.DataPropertyName = "Id";
            comboBoxColumn.ValueMember = "Id";
            comboBoxColumn.DisplayMember = "Nome";

            DataGridViewComboBoxColumn comboBoxColumn1 = new DataGridViewComboBoxColumn();
            comboBoxColumn1.DataSource = listProdutos.ToList();
            comboBoxColumn1.DataPropertyName = "Id";
            comboBoxColumn1.ValueMember = "Id";
            comboBoxColumn1.DisplayMember = "Nome";

            this.dataGridView1.Columns.Add(comboBoxColumn);
            this.dataGridView1.Columns.Add(comboBoxColumn1);
        }
        public class Produto
        {
           public int Id{get;set;}
           public string Nome{get;set;}
        }

        private void DataGridViewfrm_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'mSDNDataSet.Table' table. You can move, or remove it, as needed.
            this.tableTableAdapter.Fill(this.mSDNDataSet.Table);

        }
    }
}
