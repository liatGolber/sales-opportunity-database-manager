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
        }

        private void regB_Click(object sender, EventArgs e)
        {
            new RegisterF().ShowDialog();
        }
    }
}
