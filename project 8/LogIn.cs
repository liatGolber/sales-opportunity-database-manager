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
    public partial class LogIn : Form
    {
        public LogIn()
        {
            InitializeComponent();
        }

        private void okB_Click(object sender, EventArgs e)
        {
            User get = Program.GetUserByID(textBox1.Text);
            if (get.ID != null && get.password == textBox2.Text)
            {
                Program.currentUser = get;
                this.Close();
            }
            else
                MessageBox.Show("Invalid ID or password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {    
            if (Convert.ToInt32(e.KeyChar) - Convert.ToInt32('0') > 9)
                    e.Handled = true;
        }
    }
}
