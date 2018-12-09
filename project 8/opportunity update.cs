using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace project_8
{
    public partial class opportunity_update : Form
    {
        private Opp op;
        public opportunity_update(Opp o)
        {
            op = o;
            InitializeComponent();
            FillData();
            textBox1.Text = op.ID;
            textBox2.Text = op.name;
            textBox3.Text = op.lastN;
            textBox4.Text = op.phone;
        }
        private void FillData()
        {
            dataGridView1.Rows.Clear();
            foreach (Package p in Program.packages)
            {
                if (p.ID == op.ID)
                {
                    DataGridViewRow add = dataGridView1.Rows[0].Clone() as DataGridViewRow;
                    add.Cells[0].Value = p.lineNum;
                    add.Cells[1].Value = Program.GetPackagePrice(p.packageType).ToString() + "₪";
                    dataGridView1.Rows.Add(add);
                }
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            new addNewPackage().ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            DataGridViewRow select = dataGridView1.SelectedRows[0];
            foreach (Package p in Program.GetPackagesByID(op.ID))
                if (p.lineNum == select.Cells[0].Value.ToString())
                {
                    Program.RemovePackage(p);
                    Program.UpdatePacList();
                    FillData();
                    break;
                }
            this.Enabled = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow sent = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex];
            int n = 0;
            foreach (Package p in Program.GetPackagesByID(op.ID))
                if (p.lineNum == sent.Cells[0].Value.ToString())
                {
                    n = p.packageType;
                    break;
                }
            switch (n)
            {
                case 1:
                    {
                        richTextBox1.Text = "Intrent usage up to 100GB.\nUnlimited calls and messages.For 30 ₪/Month";
                        break;
                    }
                case 2:
                    {
                        richTextBox1.Text = "Intrent usage up to 50GB.\nUnlimited calls and messages.For 20 ₪/Month";
                        break;
                    }
                case 3:
                    {
                        richTextBox1.Text = "Intrent usage up to 10GB.\nUnlimited calls and messages.For 10 ₪/Month";
                        break;
                    }
                default:
                    {
                        richTextBox1.Text = "Please select a package.";
                        break;
                    }

            }

        }
    }
}