using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace wf1
{
    public partial class Form4 : Form
    {
        public Form4(List<double> list, string title)
        {
            InitializeComponent();
            InitializeChart(list, title);
        }

        private void InitializeChart(List<double> list, string t)
        {
            chart1.Dock = DockStyle.Fill;
            chart1.Series.Clear();
            chart1.Titles.Add(t);

            Series series = chart1.Series.Add("Ответы");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true;

            series.Points.AddXY("Присутсвует", list[0].ToString());
            series.Points.AddXY("Отсутсвует", list[1].ToString());


        }
    }
}
