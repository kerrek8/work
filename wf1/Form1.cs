using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
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
        List<double> allFobies;
        List<int> valuesTech1;
        List<int> valuesTech2;
        List<int> valuesTech3;
        List<int> valuesEst1;
        List<int> valuesEst2;
        List<int> valuesEst3;
        List<int> valuesGum1;
        List<int> valuesGum2;
        List<int> valuesGum3;

        public void Data(string path) { 
            Excel.Application ea = new Excel.Application();
            Excel.Workbook wb = ea.Workbooks.Open(path, 0, true); // открываем файл на чтение 
            
            Excel.Worksheet wsest = wb.Worksheets[3];
            Excel.Worksheet wstech = wb.Worksheets[4];
            Excel.Worksheet wsgum = wb.Worksheets[5];

            Excel.Range rangeest = wsest.Range["A2", "F58"];
            Excel.Range rangetech = wstech.Range["A2", "F156"];
            Excel.Range rangegum = wsgum.Range["A2", "F84"];

            int rowCount = rangeest.Rows.Count;
            int columnCount = rangeest.Columns.Count;

            int cntyes = 0;
            int cntno = 0;

            int allyes = 0;
            int allno = 0;

            bool flag = false;

            for (int i = 1; i <= rowCount; i++)
            {
                flag = false;
                for (int j = 1; j <= columnCount; j++)
                {
                    var cellValue = rangeest.Cells[i, j].Value2;
                    if (cellValue.ToString() == "Да")
                    {
                        allyes++;
                        flag = true;
                        break;
                    }

                }
                if (!flag) { 
                    allno++;
                }
                
            }

            for (int i = 4; i<=columnCount;i++) {
                for (int j = 1; j<=rowCount; j++) { 
                    var cellValue = rangeest.Cells[j, i].Value2;
                   
                    
                        if (cellValue.ToString() == "Да")
                        {
                            cntyes++;
                        }
                        else
                        {
                            cntno++;
                        }
                    
                    
                }
                if (i == 4)
                {
                    valuesEst1.Add(cntyes);
                    valuesEst1.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }
                else if (i == 5)
                {
                    valuesEst2.Add(cntyes);
                    valuesEst2.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }
                else if (i == 6){
                    valuesEst3.Add(cntyes);
                    valuesEst3.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }
                

            }

            rowCount = rangetech.Rows.Count;
            columnCount = rangetech.Columns.Count;

            for (int i = 1; i <= rowCount; i++)
            {
                flag = false;
                for (int j = 1; j <= columnCount; j++)
                {
                    var cellValue = rangetech.Cells[i, j].Value2;
                    if (cellValue.ToString() == "Да")
                    {
                        flag |= true;
                        allyes++;
                        break;
                    }
                    
                }
                if (!flag) { allno++; }
            }

            for (int i = 4; i <= columnCount; i++)
            {
                for (int j = 1; j <= rowCount; j++)
                {
                    var cellValue = rangetech.Cells[j, i].Value2;
                    
                    
                        if (cellValue.ToString() == "Да")
                        {
                            cntyes++;
                        }
                        else
                        {
                            cntno++;
                        }
                    
                }
                if (i == 4)
                {
                    valuesTech1.Add(cntyes);
                    valuesTech1.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }
                else if (i == 5)
                {
                    valuesTech2.Add(cntyes);
                    valuesTech2.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }
                else if ( i == 6 )
                {
                    valuesTech3.Add(cntyes);
                    valuesTech3.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }


            }

            rowCount = rangegum.Rows.Count;
            columnCount = rangegum.Columns.Count;

            for (int i = 1; i <= rowCount; i++) {
                flag = false;
                for (int j = 1; j <= columnCount; j++) {
                    var cellValue = rangegum.Cells[i, j].Value2;
                    if (cellValue.ToString() == "Да")
                    {
                        flag = true;
                        allyes++;
                        break;
                    }
                    
                }
                if (!flag) { allno++; }
            }

            for (int i = 4; i <= columnCount; i++)
            {
                for (int j = 1; j <= rowCount; j++)
                {
                    var cellValue = rangegum.Cells[j, i].Value2;
                    
                   
                        if (cellValue.ToString() == "Да")
                        {
                            cntyes++;
                        }
                        else
                        {
                            cntno++;
                        }
                    
                }
                if (i == 4)
                {
                    valuesGum1.Add(cntyes);
                    valuesGum1.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }
                else if (i == 5)
                {
                    valuesGum2.Add(cntyes);
                    valuesGum2.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }
                else if (i == 6)
                {
                    valuesGum3.Add(cntyes);
                    valuesGum3.Add(cntno);
                    cntno = 0;
                    cntyes = 0;
                }


            }

            


            allFobies[0] += allyes;
            allFobies[1] += allno;

            wb.Close(false, Type.Missing, Type.Missing);
            ea.Quit();
        }


        public Form1()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            listBox1.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // дублируется чтобы при повторном нажатии на кнопку выбора файла кнопки гасли
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            listBox1.Enabled = false;

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Выберите файл";
            ofd.Filter = "Файл с данными (*.xlsx)|*.xlsx";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                
                string Path = ofd.FileName;

                try {
                    valuesTech1 = new List<int>();
                    valuesTech2 = new List<int>();
                    valuesTech3 = new List<int>();
                    valuesEst1 = new List<int>();
                    valuesEst2 = new List<int>();
                    valuesEst3 = new List<int>();
                    valuesGum1 = new List<int>();
                    valuesGum2 = new List<int>();
                    valuesGum3 = new List<int>();
                    allFobies = new List<double>() { 0, 0 };
                    Data(Path);
                    button1.Text = "Выбранный файл: " + Path;
                    button2.Enabled = true;
                    listBox1.Enabled=true;
                    button4.Enabled = true;


                }
                catch (Exception ex) {
                    if (MessageBox.Show($"\n {ex}", "Ошибка", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) == DialogResult.Retry)
                    {
                        button1_Click(sender, e);
                    }
                }


            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<string> l = new List<string>() {
                Math.Round(100*((double)valuesTech1[0] / (valuesTech1[0] + valuesTech1[1]))).ToString(),
                Math.Round(100*((double)valuesGum1[0] / (valuesGum1[0] + valuesGum1[1]))).ToString(),
                Math.Round(100*((double)valuesEst1[0] / (valuesEst1[0] + valuesEst1[1]))).ToString()
            };
            
            Form2 form = new Form2(l);
            form.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string selectedItem = listBox1.SelectedItem as string;
            if (selectedItem == "Технические науки")
            {
                Form3 form2 = new Form3(valuesTech2, $"Считаете ли вы что у вас есть неконтролируемый страх связанный с нахождением в интернет пространстве? ({selectedItem})");
                form2.Show();
                Form3 form3 = new Form3(valuesTech3, $"Считаете ли Вы, что избавиться от интернет-фобий навсегда возможно? ({selectedItem})");
                form3.Show();
            }
            else if (selectedItem == "Гуманитарные науки")
            {
                Form3 form2 = new Form3(valuesGum2, $"Считаете ли вы что у вас есть неконтролируемый страх связанный с нахождением в интернет пространстве? ({selectedItem})");
                form2.Show();
                Form3 form3 = new Form3(valuesGum3, $"Считаете ли Вы, что избавиться от интернет-фобий навсегда возможно? ({selectedItem})");
                form3.Show();
            }
            else
            {
                Form3 form2 = new Form3(valuesEst2, $"Считаете ли вы что у вас есть неконтролируемый страх связанный с нахождением в интернет пространстве? ({selectedItem})");
                form2.Show();
                Form3 form3 = new Form3(valuesEst3, $"Считаете ли Вы, что избавиться от интернет-фобий навсегда возможно? ({selectedItem})");
                form3.Show();
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form4 form = new Form4(allFobies, "У скольки из опрошенных человек присутсвует хотя бы одна фобия связанная с нахождением в интернет пространнстве");
            form.Show();
            double firstyes = Math.Round(100 * (allFobies[0] / (allFobies[0] + allFobies[1]))); 
            List<double> percents = new List<double>() {
                firstyes,
                100-firstyes
            }; 
            Form4 form1 = new Form4(percents, "У скольки из опрошенных человек присутсвует хотя бы одна фобия связанная с нахождением в интернет пространнстве (в процентах)");
            form1.Show();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedItem = listBox1.SelectedItem as string;
            button3.Text = $"Диаграммы для: {selectedItem}";
            button3.Enabled = true;
        }
    }
}
