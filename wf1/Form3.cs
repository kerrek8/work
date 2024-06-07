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
    public partial class Form3 : Form
    {
        public Form3(List<int> list, string title)
        {
            InitializeComponent();
            InitializeChart(list, title);
        }

        private void InitializeChart(List<int> list, string title)
        {
            chart1.Dock = DockStyle.Fill;
            chart1.Series.Clear();
            chart1.Titles.Add(title);

            Series series = chart1.Series.Add("Ответы");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true;
            
            series.Points.AddXY("Да", list[0].ToString());
            series.Points.AddXY("Нет", list[1].ToString());
            

        }
    }
}
