using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace WindowsFormsApplication1
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        

        private void Form3_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.QuizTotalMark;
            textBox2.Text = Properties.Settings.Default.QuizPercMark;
            textBox3.Text = Properties.Settings.Default.Exam1TotalMark;
            textBox4.Text = Properties.Settings.Default.Exam1PercMark;
            textBox5.Text = Properties.Settings.Default.Exam2TotalMark;
            textBox6.Text = Properties.Settings.Default.Exam2PercMark;
            textBox7.Text = Properties.Settings.Default.Exam3TotalMark;
            textBox8.Text = Properties.Settings.Default.Exam3PercMark;
            textBox9.Text = Properties.Settings.Default.OutputFile;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.QuizTotalMark = textBox1.Text;
            Properties.Settings.Default.QuizPercMark = textBox2.Text;
            Properties.Settings.Default.Exam1TotalMark = textBox3.Text;
            Properties.Settings.Default.Exam1PercMark = textBox4.Text;
            Properties.Settings.Default.Exam2TotalMark = textBox5.Text;
            Properties.Settings.Default.Exam2PercMark = textBox6.Text;
            Properties.Settings.Default.Exam3TotalMark = textBox7.Text;
            Properties.Settings.Default.Exam3PercMark = textBox8.Text;
            Properties.Settings.Default.OutputFile = textBox9.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("The new values are saved");             

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog odf = new OpenFileDialog();
            odf.Title = "Output Excel File";
            odf.Filter = ".XLSX | *.xlsx";
            odf.ShowDialog();
            textBox9.Text = odf.FileName;
        }
    }
}
