﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Data;
using System.Text.RegularExpressions;

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

                Dictionary<string, double> allMarks = new Dictionary<string, double>();
                List<string> Names = new List<string>();
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
                                quizMarks.Add((double)secSheet.Cells[i, c].Value);
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
                        allMarks[cell.Value.ToString()] = avg;
                        Names.Add(secSheet.Cells[i, 2].Value.ToString());
                        mainSheet.SetValue(i, 1, Filter(cell.Value.ToString())); //setting ID
                        mainSheet.SetValue(i, 2, secSheet.Cells[i, 2].Value.ToString()); //setting Name
                        mainSheet.SetValue(i, 3, avg); //setting avg
                        //output = ((total / totalMark) * percentage) + "%";
                        double perc = (avg / totalMark) * percentage;
                        mainSheet.SetValue(i, 4, perc); //setting percentage
                    }


                }

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
                {2,7   },
                {3,9   }
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
                    for (int i = 2; i <= address_space.End.Row; i++)
                    {
                        if (secSheet.Cells[i, 3].Value != null)
                        {
                            allMarks[Filter(secSheet.Cells[i, 1].Value.ToString())] = Convert.ToDouble(secSheet.Cells[i, 3].Value);

                        }
                        else
                        {
                            allMarks[Filter(secSheet.Cells[i, 1].Value.ToString())] = 0.0d;
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

                DataTable table = new DataTable();
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
                excel.Save();


            }
        }

        private static string Filter(string input)
        {
            Regex digitsOnly = new Regex(@"[^\d]");
            return digitsOnly.Replace(input, "");
        }
    }
}
        

       