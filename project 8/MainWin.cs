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
    public partial class MainWin : Form
    {
        public MainWin()
        {
            InitializeComponent();
            helloL.Text = "Hello " + Program.currentUser.name;
            regB.Visible = Program.currentUser.isAdmin;
            FillData();
            FillReminders();
        }

        private void regB_Click(object sender, EventArgs e)
        {
            new RegisterF().ShowDialog();
        }

        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewColumn sent = dataGridView1.Columns[e.ColumnIndex];
            SearchPopUp s = new SearchPopUp(sent);
            s.Location = Cursor.Position;
            s.ShowDialog();
            if (s.value != null)
            {
                FillData(e.ColumnIndex, s.value);
                button1.Visible = true;
            }
            else
            {
                FillData();
                button1.Visible = false;
            }
            s.Dispose();
        }

        private void FillData(int colI = -1, string filter = "")
        {
            dataGridView1.Rows.Clear();
            foreach (Opp o in Program.opportunites)
            {
                DataGridViewRow add = dataGridView1.Rows[0].Clone() as DataGridViewRow;
                add.Cells[0].Value = o.ID;
                add.Cells[1].Value = o.name;
                add.Cells[2].Value = o.lastN;
                add.Cells[3].Value = o.phone;
                add.Cells[4].Value = o.status;
                add.Cells[5].Value = o.email;
                add.Cells[6].Value = o.treatedBy.ID;
                add.Cells[7].Value = o.treatedAt.ToShortDateString();
                add.Cells[8].Value = o.comment;

                switch (colI)
                {
                    case -1:
                        {
                            dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 0:
                        {
                            if (o.ID.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 1:
                        {
                            if (o.name.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 2:
                        {
                            if (o.lastN.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 3:
                        {
                            if (o.phone.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 4:
                        {
                            if (o.status.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 5:
                        {
                            if (o.email.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 6:
                        {
                            if (o.treatedBy.ID.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 7:
                        {
                            if (add.Cells[6].Value.ToString().ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 8:
                        {
                            if (o.comment.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                }
            }
        }

        private void FillReminders()
        {
            dataGridView2.Rows.Clear();
            foreach (Opp o in Program.opportunites)
            {
                DataGridViewRow add = dataGridView2.Rows[0].Clone() as DataGridViewRow;
                add.Cells[0].Value = o.ID;
                add.Cells[1].Value = o.name;
                add.Cells[2].Value = o.phone;
                int p = Program.GetStatusPrec(o.status);
                if (DateTime.Now.Date >= o.treatedAt.Date.AddDays(7).Date || p >= 80)
                    dataGridView2.Rows.Add(add);
            }
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0)
                return;
            DataGridViewRow sent = (sender as DataGridView).Rows[e.RowIndex];
            Opp o = Program.GetOpByID(sent.Cells[0].Value.ToString());
            this.Hide();
            new opportunity_page(o).ShowDialog();
            FillData();
            FillReminders();
            this.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FillData();
            button1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            new opportunity_page().ShowDialog();
            FillData();
            FillReminders();
            this.Show();
        }
    }
}
