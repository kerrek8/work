using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace wf1
{
    public partial class Form1 : Form
    {
        string[][] data;

        public void Data(string path) { 
            Excel.Application ea = new Excel.Application();
            Excel.Workbook wb = ea.Workbooks.Open(path, 0, true); // открываем файл на чтение 
            Excel.Worksheet ws = wb.Worksheets[1]; // выбираем первую страницу (лист)
            Excel.Range range = ws.Range["A1", "R" + ws.UsedRange.Rows.Count];
            int rowCount = range.Rows.Count;
            int columnCount = range.Columns.Count;
            string[,] datas = new string[rowCount, columnCount];
            for (int i = 2; i<rowCount; i++) { 
                for (int j = 2; j<columnCount; j++)
                {
                    var cellValue = range.Cells[i, j].Value2;
                    if (cellValue != null)
                    {
                        datas[i-2, j-2] = cellValue.ToString();
                    }
                    else
                    {
                        datas[i-2, j-2] = "";
                    }
                    
                }
            }


            MessageBox.Show(datas[0,0]);
            
            wb.Close(false, Type.Missing, Type.Missing);
            ea.Quit();
        }


        public Form1()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Выберите файл";
            ofd.Filter = "Файл с данными (*.xlsx)|*.xlsx";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                
                string Path = ofd.FileName;

                try { 
                    Data(Path);
                    button1.Text = "Выбранный файл: " + Path;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;


                }
                catch (Exception ex) {
                    if (MessageBox.Show("Неподходящий файл, проверьте данные в файле!\n" + "Выбранный файл: " + Path + $"\n {ex}", "Ошибка", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) == DialogResult.Retry)
                    {
                        button1_Click(sender, e);
                    }
                }


            }

        }

            private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
    }
}
