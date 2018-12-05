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
    public partial class Worker : Form
    {
        public Worker()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.chart1.Series["Average status"].Points.AddXY("add name of salesman",5);
            this.chart1.Series["Average status"].Points.AddXY("everyone", 15);
            this.chart1.Series["Deals won"].Points.AddXY("add name of salesman", 5);
            this.chart1.Series["Deals won"].Points.AddXY("everyone", 15);
            this.chart1.Series["Deals lost"].Points.AddXY("add name of salesman", 5);
            this.chart1.Series["Deals lost"].Points.AddXY("everyone", 15);
            this.chart1.Series["Total deals"].Points.AddXY("add name of salesman", 5);
            this.chart1.Series["Total deals"].Points.AddXY("everyone", 15);
            this.chart1.Series["Average price"].Points.AddXY("add name of salesman", 5);
            this.chart1.Series["Average price"].Points.AddXY("everyone", 15);
        }
    }
}
