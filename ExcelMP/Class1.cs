using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.Threading;
namespace ExcelMP
{
    public class ExMP
    {
       


        public static void AssessQuiz(string path1, string path2, int sheet1, int sheet2, int totalMark, int percentage)
        {
            FileInfo mainFile = new FileInfo(path1);
            FileInfo secondaryFile = new FileInfo(path2);
            if (mainFile.Exists == true && secondaryFile.Exists == true)
            {
                ExcelPackage mainPackage = new ExcelPackage(mainFile); //original file for edit
                ExcelWorksheet mainSheet = mainPackage.Workbook.Worksheets[sheet1];

               
               
                using (ExcelPackage secPackage = new ExcelPackage(secondaryFile))
                {
                    ExcelWorksheet secSheet = secPackage.Workbook.Worksheets[sheet2];
                    var address_space2 = secSheet.Dimension;
                    
                    for (int i = 2; i <= address_space2.End.Row; i++) //row
                    {
                        List<double> quizMarks = new List<double>();
                        for (int c = 3; c <= address_space2.End.Column; c++) //columns
                        {
                            if (secSheet.Cells[i, c].Value != null)
                            {
                                string value = Filter(secSheet.Cells[i, c].Value.ToString()); //removing unnecessary characters from them excel file
                                if (value != "")
                                {
                                    quizMarks.Add(Convert.ToDouble(value));
                                }
                                else
                                {
                                    quizMarks.Add(0);
                                }

                            }
                            else
                            {
                                quizMarks.Add(0);
                            }


                        }
                        //finishing looping through the columns

                        var cell = secSheet.Cells[i, 1];

                        var output = from mark in quizMarks
                                     orderby mark descending
                                     select mark;
                        double avg = output.Take(5).Average();
                        //allMarks[cell.Value.ToString()] = avg;
                        //Names.Add(secSheet.Cells[i, 2].Value.ToString());
                        if (cell.Value != null)
                        {


                            mainSheet.SetValue(i, 1, Filter(cell.Value.ToString())); //setting ID
                            mainSheet.SetValue(i, 2, secSheet.Cells[i, 2].Value.ToString()); //setting Name
                            mainSheet.SetValue(i, 3, avg); //setting avg
                            //output = ((total / totalMark) * percentage) + "%";
                            double perc = (avg / totalMark) * percentage;
                            mainSheet.SetValue(i, 4, perc); //setting percentage
                        }

                    }


                }
                //mainPackage.Dispose();
                mainPackage.Save();









                //ExcelWorksheet work_sheet = workbook.Worksheets[sheet];
                //var address_space = work_sheet.Dimension;

                //var rows = work_sheet.Cells[1, 1, address_space.End.Row, 1].ToList();
                //foreach (var row in rows)
                //{
                //    if (row.Text == "s2612269")
                //    {
                //        var tmp1 = row.Start;


                //        work_sheet.SetValue(tmp1.Row, tmp1.Column + 1,10);
                //        //var tmp2 = row.Address;
                //        //var tmp3 = row.End;
                //        //var tmp4 = row.FullAddress;
                //    }
                //}
                //file_Package.Save();

                //work_sheet.Cells["A1"].LoadFromDataTable(,);

            }



        }

        public static void AssesExam(string path1, string path2, int sheet1, int sheet2, int totalMark, int percentage, int examNo)
        {
            Dictionary<int, int> ExamColumnMapping = new Dictionary<int, int>() 
            { 
                {1 , 5 },
                {2,  7 },
                {3,  9 }
            };

            FileInfo mainFile = new FileInfo(path1);
            FileInfo secondaryFile = new FileInfo(path2);
            List<ExcelRangeBase> Ids = new List<ExcelRangeBase>();
            Dictionary<string, double> allMarks = new Dictionary<string, double>();

            if (mainFile.Exists == true && secondaryFile.Exists == true)
            {
                using (ExcelPackage secPackage = new ExcelPackage(secondaryFile))
                {
                    ExcelWorksheet secSheet = secPackage.Workbook.Worksheets[sheet2];
                    var address_space = secSheet.Dimension;
                    for (int i = 2; i <= address_space.End.Row; i++) //loop through rows starting with
                    {
                        if (secSheet.Cells[i, 3].Value != null )
                        {
                            if ((secSheet.Cells[i, 1].Value != null))
                            {
                                  allMarks[Filter(secSheet.Cells[i, 1].Value.ToString())] = Convert.ToDouble(Filter(secSheet.Cells[i, 3].Value.ToString())); //reading Ids and Marks ???
                                  string value = Filter(secSheet.Cells[i, 3].Value.ToString());
                                  if (value != "")
                                  {
                                       allMarks[Filter(secSheet.Cells[i, 1].Value.ToString())] = Convert.ToDouble(Filter(value));
                                  }
                                  else
                                  {
                                     allMarks[Filter(secSheet.Cells[i, 1].Value.ToString())] = 0.0d; //setting the value to be red to zero

                                  }
                            }

                        }
                        else
                        {
                            if ((secSheet.Cells[i, 1].Value != null))
                            {
                                allMarks[Filter(secSheet.Cells[i, 1].Value.ToString())] = 0.0d;
                            }
                           
                        }
                    }
                }

                using (ExcelPackage mainPackage = new ExcelPackage(mainFile))
                {
                    ExcelWorksheet mainSheet = mainPackage.Workbook.Worksheets[sheet1];
                    var address_space = mainSheet.Dimension;
                    Ids = mainSheet.Cells[2, 1, address_space.End.Row, 1].ToList(); ; //list of all Ids
                    foreach (var Id in Ids)
                    {
                        string tempId = Id.Text;

                        if (allMarks.ContainsKey(tempId))
                        {
                            double mark = allMarks[tempId];
                            mainSheet.SetValue(Id.End.Row, ExamColumnMapping[examNo], mark); //setting Mark
                            double perc = (mark / totalMark) * percentage;
                            mainSheet.SetValue(Id.End.Row, ExamColumnMapping[examNo] + 1, perc); //setting Percentage
                        }
                    }
                    mainPackage.Save();
                }







            }
        }


        public static void TotalPercentage(string path1, int sheet)
        {
            FileInfo mainFile = new FileInfo(path1);
            if (mainFile.Exists == true)
            {

                using (ExcelPackage mainPackage = new ExcelPackage(mainFile))
                {
                    ExcelWorksheet Sheet = mainPackage.Workbook.Worksheets[sheet];
                    var address_space = Sheet.Dimension;
                    var Ids = Sheet.Cells[2, 1, address_space.End.Row, 1].ToList(); ; //list of all Ids
                    for (int i = 2; i <= address_space.End.Row; i++)
                    {
                        double quiz = 0;
                        double exam1 = 0;
                        double exam2 = 0;
                        double exam3 = 0;
                        if (Sheet.Cells[i, 4].Value != null && Sheet.Cells[i, 4].Text != "")
                        {
                            quiz = Convert.ToDouble(Sheet.Cells[i, 4].Text);
                        }

                        if (Sheet.Cells[i, 6].Value != null && Sheet.Cells[i, 6].Text != "")
                        {
                            exam1 = Convert.ToDouble(Sheet.Cells[i, 6].Text);
                        }

                        if (Sheet.Cells[i, 8].Value != null && Sheet.Cells[i, 8].Text != "")
                        {
                            exam2 = Convert.ToDouble(Sheet.Cells[i, 8].Text);
                        }

                        if (Sheet.Cells[i, 10].Value != null && Sheet.Cells[i, 10].Text != "")
                        {
                            exam3 = Convert.ToDouble(Sheet.Cells[i, 10].Text);
                        }

                        Sheet.SetValue(i, 11, (quiz + exam1 + exam2 + exam3));

                    }
                    mainPackage.Save();
                }




            }
        }


        public static void CreateFile(string path)
        {



            if (!File.Exists(path))
            {
                var excel = new ExcelPackage(new FileInfo(path));
                var ws = excel.Workbook.Worksheets.Add("Sheet1");

                System.Data.DataTable table = new System.Data.DataTable();
                table.Columns.Add("Id", typeof(string));
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Quiz Average", typeof(string));
                table.Columns.Add("Quiz Percentage", typeof(string));
                table.Columns.Add("Exam 1", typeof(string));
                table.Columns.Add("Exam 1 Percentage", typeof(string));
                table.Columns.Add("Exam 2", typeof(string));
                table.Columns.Add("Exam 2 Percentage", typeof(string));
                table.Columns.Add("Exam 3", typeof(string));
                table.Columns.Add("Exam 3 Percentage", typeof(string));
                table.Columns.Add("Total Percentage", typeof(string));
                ws.Cells.LoadFromDataTable(table, true);

                /*Color Settings*/
                ws.Column(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Column(1).Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                ws.Column(3).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Column(3).Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                ws.Column(5).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Column(5).Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                ws.Column(7).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Column(7).Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                ws.Column(9).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Column(9).Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                ws.Column(11).Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Column(11).Style.Fill.BackgroundColor.SetColor(Color.LightCyan);


                /**/
                ws.Column(1).Width = 12.5d;
                ws.Column(2).Width = 36.5d;
                ws.Column(3).Width = 13d;
                ws.Column(4).Width = 16d;
                ws.Column(5).Width = 8d;
                ws.Column(6).Width = 18d;
                ws.Column(7).Width = 8d;
                ws.Column(8).Width = 18d;
                ws.Column(8).Width = 18d;
                ws.Column(9).Width = 8d;
                ws.Column(10).Width = 18d;
                ws.Column(11).Width = 18d;
                excel.Save();


            }
        }

        private static string Filter(string input)
        {
            Regex digitsOnly = new Regex(@"[^\d\.,]");
            return digitsOnly.Replace(input, "");
        }

        public static void OpenXLSXFile(string filePath)
        {



            if (File.Exists(filePath))
            {

                try
                {
                    Application excel = new Application();
                    excel.Visible = true;
                    Workbook wb = excel.Workbooks.Open(filePath);

                    excel.Width = 400d;
                    wb.Close();
                    excel.Quit();
                }
                catch (Exception)
                {
                    
                    //do nothing
                }
                
               



            }
            else
            {
                throw new FileNotFoundException("File was not found");
            }

         
           
         

           
            
            

        }

        //private static void SetNewCulture()
        //{
        //    oldCulture = Thread.CurrentThread.CurrentCulture;
        //    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

        //}

        //private static void ReSetOldCulture() 
        //{
        //    Thread.CurrentThread.CurrentCulture = oldCulture;
        //}



       
    }
}
        

       