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

        private void button1_Click(object sender, EventArgs e)
        {
            User omer = Program.GetUserByID("208063511");
            float[] omerstats = Program.GetStatistics(omer);
            float[] everyoneStats = Program.GetStatistics();
            this.chart1.Series["worker"].Name = omer.name;
            this.chart1.Series[omer.name].Points.AddXY("Status average", (int)omerstats[0]);
            this.chart1.Series[omer.name].Points.AddXY("Deals won", omerstats[1]);
            this.chart1.Series[omer.name].Points.AddXY("Deals lost", omerstats[2]);
            this.chart1.Series[omer.name].Points.AddXY("Number of opportunites", omerstats[3]);
            this.chart1.Series[omer.name].Points.AddXY("Average price of sales", (int)omerstats[4]);

            this.chart1.Series["overall"].Points.AddXY("Status average", (int)everyoneStats[0]);
            this.chart1.Series["overall"].Points.AddXY("Deals won", everyoneStats[1]);
            this.chart1.Series["overall"].Points.AddXY("Deals lost", everyoneStats[2]);
            this.chart1.Series["overall"].Points.AddXY("Number of opportunites", everyoneStats[3]);
            this.chart1.Series["overall"].Points.AddXY("Average price of sales", (int)everyoneStats[4]);


            //this.chart1.Series["Deals won"].Points.AddXY(omer.name, omerstats[1]);
            //this.chart1.Series["Deals won"].Points.AddXY("Overall", everyoneStats[1]);

            //this.chart1.Series["Deals lost"].Points.AddXY(omer.name, omerstats[2]);
            //this.chart1.Series["Deals lost"].Points.AddXY("Overall", everyoneStats[2]);

            //this.chart1.Series["Number of opportunites"].Points.AddXY(omer.name, omerstats[3]);
            //this.chart1.Series["Number of opportunites"].Points.AddXY("Overall", everyoneStats[3]);

            //this.chart1.Series["Average price of sales"].Points.AddXY(omer.name, omerstats[4]);
            //this.chart1.Series["Average price of sales"].Points.AddXY("Overall", everyoneStats[4]);
        }
    }
}
