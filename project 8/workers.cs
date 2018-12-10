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
    public partial class workers : Form
    {
        public workers()
        {
            InitializeComponent();
            FillData();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            DataGridViewRow sent = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex];
            User u = Program.GetUserByID(sent.Cells[0].Value != null ? sent.Cells[0].Value.ToString() : "");
            if (u.ID == null)
                return;
            float[] uStats = Program.GetStatistics(u);
            float[] overStats = Program.GetStatistics();
            label1.Text = u.name + " won:" + uStats[1] / 100 * overStats[1];
            label2.Text = "Overall won:" + overStats[1];
            label3.Text = u.name + " lost:" + uStats[2] / 100 * overStats[2];
            label4.Text= "Overall lost:" + overStats[2];

            chart1.Series["ser"].Points.Clear();
            chart1.Series["ser"].Points.AddXY(u.name, uStats[1] / 100);
            chart1.Series["ser"].Points.AddXY("Overall", 1 - (uStats[1] / 100));
            chart1.Series["ser2"].Points.Clear();
            chart1.Series["ser2"].Points.AddXY(u.name, uStats[2] / 100);
            chart1.Series["ser2"].Points.AddXY("Overall", 1 - (uStats[2] / 100));
            chart2.Series[0].Points.Clear();
            chart2.Series["Overall"].Points.Clear();
            chart2.Series[0].Name = u.name;
            chart2.Series[u.name].Points.AddXY("Deals Count", uStats[3]);
            chart2.Series["Overall"].Points.AddXY("Deals Count", overStats[3]);
            chart2.Series[u.name].Points.AddXY("status avg", (int)uStats[0]);
            chart2.Series["Overall"].Points.AddXY("status avg", (int)overStats[0]);
            chart2.Series[u.name].Points.AddXY("price avg", (int)uStats[4]);
            chart2.Series["Overall"].Points.AddXY("price avg", (int)overStats[4]);
        }

        private void FillData()
        {
            dataGridView1.Rows.Clear();
            foreach (User u in Program.userList)
            {
                DataGridViewRow add = dataGridView1.Rows[0].Clone() as DataGridViewRow;
                add.Cells[0].Value = u.ID;
                add.Cells[1].Value = u.name;
                add.Cells[2].Value = u.lastN;
                dataGridView1.Rows.Add(add);

            }

        }

        private void regB_Click(object sender, EventArgs e)
        {
            new RegisterF().ShowDialog();
        }

        private void workers_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
