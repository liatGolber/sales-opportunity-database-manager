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
            textBox7.Text = op.status;
            textBox9.Text = op.treatedAt.ToShortDateString();
            richTextBox1.Text = op.comment;

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}
