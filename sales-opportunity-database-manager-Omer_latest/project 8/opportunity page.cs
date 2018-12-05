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
            updatedTextBoxes();

        }
        public opportunity_page()
        {
            InitializeComponent();
            textBox3.ReadOnly = false;
            textBox7.Text = "New";
            textBox8.Text = "10%";
            textBox9.Text = DateTime.Now.Date.ToShortDateString();
        }

        private void updatedTextBoxes()
        {
            textBox1.Text = opp.name;
            textBox2.Text = opp.lastN;
            textBox3.Text = opp.ID;
            textBox4.Text = opp.phone;
            textBox5.Text = opp.email;
            textBox7.Text = opp.status.Substring(0, opp.status.IndexOf('('));
            textBox8.Text = opp.status.Substring(opp.status.IndexOf('(') + 1, opp.status.IndexOf(')') - 1 - opp.status.IndexOf('('));
            textBox9.Text = opp.treatedAt.ToShortDateString();
            richTextBox1.Text = opp.comment;
            button2.Visible = button1.Visible = false;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            new opportunity_update(opp).ShowDialog();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            button1.Visible = true;
            if (opp.ID != null)
                button2.Visible = true;
            if (opp.ID != null && opp.ID == textBox3.Text && opp.name == textBox1.Text && opp.lastN == textBox2.Text
                && opp.phone == textBox4.Text && textBox5.Text == opp.email && richTextBox1.Text == opp.comment)
            {
                button2.Visible = button1.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            updatedTextBoxes();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "")
            {
                Program.InsertUpdateOpp(textBox3.Text, textBox1.Text, textBox2.Text, textBox4.Text, textBox5.Text, DateTime.Now, textBox7.Text + "(" + textBox8.Text + ")", Program.currentUser.ID, richTextBox1.Text);
                Program.UpdateOppList();
                opp = Program.GetOpByID(textBox3.Text);
                updatedTextBoxes();
                textBox3.ReadOnly = true;
            }
            else
                MessageBox.Show("Please fill all the required fields.", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }
    }
}
