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
    public partial class RegisterF : Form
    {
        public RegisterF()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string err = "";
            if (textBox1.Text.Length != 9)
                err += "ID must be 9 digits long\n";
            if (textBox2.Text == "")
                err += "Please enter a name\n";
            if (textBox3.Text == "")
                err += "Please enter a last name\n";
            if (textBox5.Text == "")
                err += "Please enter a password\n";
            if (err != "")
                MessageBox.Show(err, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                Cursor.Current = Cursors.WaitCursor;
                if (Program.GetUserByID(textBox1.Text).ID == null)
                {
                    Program.InsertNewUser(textBox1.Text, textBox2.Text, textBox3.Text, textBox5.Text, checkBox1.Checked);
                    this.Close();
                }
                else
                    MessageBox.Show(textBox1.Text + " is already registered please use a differnt ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor.Current = Cursors.Default;

            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Convert.ToInt32(e.KeyChar) - Convert.ToInt32('0') > 9)
                e.Handled = true;
        }
    }
}
