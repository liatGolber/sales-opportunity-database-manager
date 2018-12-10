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
        public Opp opp;
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
            button1.Text = "Add";
            button4.Visible = false;
        }

        private void updatedTextBoxes()
        {
            textBox1.Text = opp.name;
            textBox2.Text = opp.lastN;
            textBox3.Text = opp.ID;
            textBox4.Text = opp.phone;
            textBox5.Text = opp.email;
            textBox7.Text = opp.status.Substring(0, opp.status.IndexOf('('));
            textBox8.Text = Program.GetStatusPrec(opp.status).ToString();
            textBox9.Text = opp.treatedAt.ToShortDateString();
            richTextBox1.Text = opp.comment;
            button2.Visible = button1.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            opportunity_update ou = new opportunity_update(opp);
            DialogResult d = ou.ShowDialog();
            opp = ou.op; 
            updatedTextBoxes();
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
            this.Enabled = false;
            string err = "";
            if (textBox1.Text == "")
                err += "Please enter a name\n";
            if (textBox2.Text == "")
                err += "Please enter a last name\n";
            if (textBox3.Text.Length != 9)
                err += "ID must be 9 digits long\n";
            if (textBox4.Text.Length != 10)
                err += "Phone must be 10 digits long\n";
            if (textBox5.Text == "")
                err += "Please enter a email\n";
            if (err != "")
                MessageBox.Show(err, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                if (opp.ID == null && Program.GetOpByID(textBox3.Text).ID == null)
                {
                    Program.InsertUpdateOpp(textBox3.Text, textBox1.Text, textBox2.Text, textBox4.Text, textBox5.Text, DateTime.Now, textBox7.Text + "(" + textBox8.Text + ")", Program.currentUser.ID, richTextBox1.Text);
                    button1.Text = "Update";
                    button4.Visible = true;
                }
                else if (opp.ID != null)
                    Program.InsertUpdateOpp(textBox3.Text, textBox1.Text, textBox2.Text, textBox4.Text, textBox5.Text, DateTime.Now, textBox7.Text + "(" + textBox8.Text + ")", Program.currentUser.ID, richTextBox1.Text);
                else
                    MessageBox.Show("ID already used.");
                Program.UpdateOppList();
                opp = Program.GetOpByID(textBox3.Text);
                updatedTextBoxes();
                textBox3.ReadOnly = true;
            }
            this.Enabled = true;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Convert.ToInt32(e.KeyChar) - Convert.ToInt32('0') > 9)
                e.Handled = true;
        }


    }
}


