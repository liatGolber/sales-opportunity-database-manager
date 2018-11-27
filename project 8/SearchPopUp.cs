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
    public partial class SearchPopUp : Form
    {
        public string value;
        public SearchPopUp(DataGridViewColumn col)
        {
            InitializeComponent();
            label1.Text += " " + col.HeaderText;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
            {
                if (textBox1.Text != "")
                    value = textBox1.Text;
                this.Close();
            }
        }
    }
}
