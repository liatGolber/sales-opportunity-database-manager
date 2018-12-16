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
    public partial class addNewPackage : Form
    {
        Opp opp;
        public addNewPackage(Opp o)
        {
            InitializeComponent();
            comboBox3.SelectedIndex = 0;
            opp = o;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox3.SelectedIndex)
            {
                case 0:
                    {
                        richTextBox1.Text = "Intrent usage up to 100GB.\nUnlimited calls and messages.For 30 ₪/Month";
                        break;
                    }
                case 1:
                    {
                        richTextBox1.Text = "Intrent usage up to 50GB.\nUnlimited calls and messages.For 20 ₪/Month";
                        break;
                    }
                case 2:
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

        private void button1_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            string err = "";
            if (textBox4.Text == "")
                err += "Please enter a phone number\n";
            else if (Program.PhoneUsed(textBox4.Text))
                err += "Phone already taken\n";
            if (textBox4.Text.Length != 10)
                err += "Phone Num Must Heve 10 Digits\n";
                if (err != "")
                MessageBox.Show(err, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                Program.InsertUpdatePackage(opp.ID, textBox4.Text, Convert.ToInt32(comboBox3.Text), false);
                this.Close();
            }
            this.Enabled = true;
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Convert.ToInt32(e.KeyChar) - Convert.ToInt32('0') > 9)
                e.Handled = true;
        }
    }

}

