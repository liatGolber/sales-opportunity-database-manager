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
    public partial class opportunity_page : Form
    {
        private Opp opp;
        public opportunity_page(Opp op)
        {
            InitializeComponent();
            opp = op;
            textBox1.Text = op.name;
            textBox2.Text = op.lastN;
            textBox3.Text = op.ID;
            textBox4.Text = op.phone;
            textBox5.Text = op.email;
            textBox7.Text = op.status.Substring(0, op.status.IndexOf('('));
            textBox8.Text = op.status.Substring(op.status.IndexOf('(') + 1, op.status.IndexOf(')') - 1 - op.status.IndexOf('('));
            textBox9.Text = op.treatedAt.ToShortDateString();
            richTextBox1.Text = op.comment;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            new opportunity_update(opp).ShowDialog();
        }
    }
}
