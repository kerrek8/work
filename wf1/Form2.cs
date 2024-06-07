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
    public partial class Form2 : Form
    {
        public Form2(List<string> list)
        {
            InitializeComponent();
            InitializeChart(list);
        }

        private void InitializeChart(List<string> list)
        {
            chart1.Dock = DockStyle.Fill;
            chart1.Series.Clear();
            chart1.Titles.Add("Процент ответивших людей которые боятся знакомится в интернете");

            Series series = chart1.Series.Add("Проценты");
            series.ChartType = SeriesChartType.Column;
            series.IsValueShownAsLabel = true;


            
            series.Points.AddXY("Технические науки", list[0]);
            series.Points.AddXY("Гуманитарные науки", list[1]);
            series.Points.AddXY("Естественные науки", list[2]);



            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisX.Title = "";
            chart1.ChartAreas[0].AxisY.Title = "";
            ChartArea chartArea = chart1.ChartAreas[0];
            chartArea.AxisY.Maximum = 100; // максимальное значение
            // Скрыть сетку
            chartArea.AxisX.MajorGrid.Enabled = false;
            chartArea.AxisX.MinorGrid.Enabled = false;
            chartArea.AxisY.MajorGrid.Enabled = true;
            chartArea.AxisY.MinorGrid.Enabled = false;

            // Настройка цвета столбцов
            series.Points[0].Color = System.Drawing.Color.Green;
            series.Points[1].Color = System.Drawing.Color.Blue;
            series.Points[2].Color = System.Drawing.Color.Purple;
            
            chart1.Legends[0].Enabled = false;
        }

    }
}
