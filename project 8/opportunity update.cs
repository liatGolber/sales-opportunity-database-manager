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
        }
        private void FillData()
        {
            dataGridView1.Rows.Clear();
            foreach (Package p in Program.onPackages)
            {
                if (p.ID == op.ID)
                {
                    DataGridViewRow add = dataGridView1.Rows[0].Clone() as DataGridViewRow;
                    add.Cells[0].Value = p.lineNum;
                    add.Cells[1].Value = p.startD.ToShortDateString();
                    add.Cells[2].Value = p.endD.ToShortDateString();
                    dataGridView1.Rows.Add(add);
                }
            }

        }
    }
}
