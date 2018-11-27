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
            regB.Enabled = Program.currentUser.isAdmin;
            FillData();
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
                FillData(e.ColumnIndex, s.value);
            else
                FillData();
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
                add.Cells[3].Value = o.Phone;
                add.Cells[4].Value = o.status;
                add.Cells[5].Value = o.treatedBy.ID;
                add.Cells[6].Value = o.treatedAt.ToShortDateString();
                add.Cells[7].Value = o.comment;

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
                            if (o.Phone.ToUpper().Contains(filter.ToUpper()))
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
                            if (o.treatedBy.ID.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 6:
                        {
                            if (add.Cells[6].Value.ToString().ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                    case 7:
                        {
                            if (o.comment.ToUpper().Contains(filter.ToUpper()))
                                dataGridView1.Rows.Add(add);
                            break;
                        }
                }
            }
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridViewRow sent = dataGridView1.Rows[e.RowIndex];
            Opp o = Program.GetOpByID(sent.Cells[0].Value.ToString());
            this.Hide();
            new opportunity_page(o).ShowDialog();
            this.Show();
        }
    }
}
