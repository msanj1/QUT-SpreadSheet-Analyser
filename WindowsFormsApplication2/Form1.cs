using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using SpreadSheetExcel;
using System.Threading;
using ExcelMP;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
       
        string comp;
        public Form1()
        {
          
            Thread t = new Thread(new ThreadStart(splashscreen));
            t.Start();
            Thread.Sleep(5000);

            InitializeComponent();
            t.Abort();
            
            textBox4.Text = "1";
            ExcelMP.ExMP.CreateFile(Properties.Settings.Default.OutputFile);

        }

        public void splashscreen() {
            Application.Run(new Form2());
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        

        private void button2_Click(object sender, EventArgs e)
        {
            string temp = textBox2.Text;
          
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Open the excel file that is used for mark comparison";
            ofd.Filter = ".XLSX | *.xlsx";
            ofd.ShowDialog();
            //string name =  ofd.;
            comp = ofd.FileName;
            if (temp != "" && comp == "")
            {
                textBox2.Text = temp;
            }
            else
            {
                textBox2.Text = comp;
            }
          
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int inputSheetNo = 1;

            int compSheetNo = Convert.ToInt32(textBox4.Text);
            string input = Properties.Settings.Default.OutputFile;
            try
            {
                if (radioButton1.Checked == true && textBox2.Text != "")
                {
                    ExMP.AssessQuiz(input, comp, inputSheetNo, compSheetNo, Convert.ToInt32(Properties.Settings.Default.QuizTotalMark), Convert.ToInt32(Properties.Settings.Default.QuizPercMark));
                    
                }
                else if (radioButton9.Checked == true && textBox2.Text != "")
                {
                    ExMP.AssesExam(input, comp, inputSheetNo, compSheetNo, Convert.ToInt32(Properties.Settings.Default.Exam1TotalMark), Convert.ToInt32(Properties.Settings.Default.Exam1PercMark),1);
                }
                else if (radioButton10.Checked == true && textBox2.Text != "")
                {
                    ExMP.AssesExam(input, comp, inputSheetNo, compSheetNo, Convert.ToInt32(Properties.Settings.Default.Exam2TotalMark), Convert.ToInt32(Properties.Settings.Default.Exam2PercMark),2);
                }
                else if (radioButton11.Checked == true && textBox2.Text != "")
                {
                    ExMP.AssesExam(input, comp, inputSheetNo, compSheetNo, Convert.ToInt32(Properties.Settings.Default.Exam3TotalMark), Convert.ToInt32(Properties.Settings.Default.Exam3PercMark),3);
                }
                else
                {
                    MessageBox.Show("Please Make sure to specify an Input in Excel 2007 or later format");
                }
                ExMP.TotalPercentage(input, inputSheetNo);
                MessageBox.Show("Finished Editing Files");
                ExMP.OpenXLSXFile(input);
            }
            catch (Exception error)
            {
                 MessageBox.Show(error.Message);
            }
            //ExMP.AssessQuiz(input, comp, inputSheetNo, compSheetNo, Convert.ToInt32(Properties.Settings.Default.QuizTotalMark), Convert.ToInt32(Properties.Settings.Default.QuizPercMark));
            //ExMP.AssesExam(input, comp, inputSheetNo, compSheetNo, Convert.ToInt32(Properties.Settings.Default.QuizTotalMark), Convert.ToInt32(Properties.Settings.Default.QuizPercMark),1);
           
        }

        //private void button3_Click(object sender, EventArgs e)
        //{

        //    try
        //    {
        //        int inputSheetNo = Convert.ToInt32(textBox3.Text);

        //        int compSheetNo = Convert.ToInt32(textBox4.Text);
        //        //string totalMark = textBox5.Text;
        //        //string percentageMark = textBox6.Text;

        //        if (inputSheetNo <= 8 && inputSheetNo >= 0 && compSheetNo <= 8 && compSheetNo >= 0 && input != "" && comp != "")
        //        {
        //            --inputSheetNo;
        //            --compSheetNo;
        //            if (radioButton1.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
                        
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                //DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz1");
                       
        //                DataTable output =  ExcelReader.CompareExcelFiles(table1, table2, "Quiz1");
                       
                       
        //                 ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));
                        

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton2.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz2");
        //                ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));


        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton3.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz3");
        //                ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton4.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz4");
        //                ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton5.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz5");
        //                ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton6.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz6");
        //                ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton7.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz7");
        //                ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton8.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Quiz8");
        //                ExcelReader.CalculatePercentage(output, "quiz", Convert.ToDouble(Properties.Settings.Default.QuizPercMark), Convert.ToDouble(Properties.Settings.Default.QuizTotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton9.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Exam1");
        //                ExcelReader.CalculatePercentage(output, "exam1", Convert.ToDouble(Properties.Settings.Default.Exam1PercMark), Convert.ToDouble(Properties.Settings.Default.Exam1TotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton10.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Exam2");
        //                ExcelReader.CalculatePercentage(output, "exam2", Convert.ToDouble(Properties.Settings.Default.Exam2PercMark), Convert.ToDouble(Properties.Settings.Default.Exam2TotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            if (radioButton11.Checked)
        //            {
        //                DataTable table1 = ExcelReader.ExtracttoDataTable(input, inputSheetNo);
        //                DataTable table2 = ExcelReader.ExtracttoDataTable(comp, compSheetNo);
        //                DataTable output = ExcelReader.CompareExcelFiles(table1, table2, "Exam3");
        //                ExcelReader.CalculatePercentage(output, "exam3", Convert.ToDouble(Properties.Settings.Default.Exam3PercMark), Convert.ToDouble(Properties.Settings.Default.Exam3TotalMark));

        //                ExcelReader.ExportToXLSX(input, output, "Sheet1");
        //            }
        //            MessageBox.Show("Finished changing the file");
        //            textBox1.Clear();
        //            textBox2.Clear();
        //            textBox3.Clear();
        //            textBox4.Clear();
                   
        //        }
        //        else
        //        {
        //            MessageBox.Show("Could not make any changes!!!! Please check your input and try again");
        //        }
        //    }
        //    catch (Exception error)
        //    {

        //        MessageBox.Show(error.Message);
        //    }

          
          
        //}

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.ShowDialog();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

       

      

        

       
    }
}
