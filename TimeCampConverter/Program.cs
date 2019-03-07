using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace TimecampConverter
{
    class Program
    {
        private static string path = "C:\\Users\\coren\\Desktop\\TimeCamp\\";
        private static string inputFileName = "Reports   TimeCamp.html";
        private static string outputFileName = "Report_Mars.xlsx";

        static void Main(string[] args)
        {
            Console.WriteLine("TIMECAMP CONVERT");
            Console.WriteLine("Reading File");
            var filePath = path + inputFileName;
            //TODO controller les max (En heure (max 50h/s ou 140h/m))
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException();
            }
            List<UniteeTask> uniteeTaskList = new List<UniteeTask>();
            UniteeTaskComparator comparator = new UniteeTaskComparator();
            FillUniteeTaskList(uniteeTaskList);
            uniteeTaskList = MergeDoubleTaskByDay(uniteeTaskList);
            uniteeTaskList.Sort(comparator);

            //Création du fichier Excel
            // create xls if not exists
            if (!File.Exists(path+outputFileName))
            {
                XSSFWorkbook wb;
                XSSFSheet sh;
                wb = new XSSFWorkbook();

                // create sheet
                sh = (XSSFSheet)wb.CreateSheet("Sheet1");
                for (int i = 0; i < uniteeTaskList.Count; i++)
                {
                    var task = uniteeTaskList[i];
                    var r = sh.CreateRow(i);
                    //Date 
                    var cel = r.CreateCell(0);
                    cel.SetCellValue(task.Date.ToString());
                    //Temps facturé 
                    var cell = r.CreateCell(1);
                    double value = (double)(task.TimeSpent / 60);
                    cell.SetCellValue(value);
                    
                    //Projet
                    r.CreateCell(2).SetCellValue(task.Project);
                    //Type    
                    r.CreateCell(3).SetCellValue("Developpement");
                    //Réf facture
                    r.CreateCell(4).SetCellValue("");
                    //Description
                    r.CreateCell(5).SetCellValue(task.Description);
                }

                using (var fs = new FileStream(path+outputFileName, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(fs);
                }
            }

            Console.ReadKey();
        }

        private static List<UniteeTask> MergeDoubleTaskByDay(List<UniteeTask> uniteeTaskList)
        {
            List<UniteeTask> mergeUniteeTaskList = new List<UniteeTask>();

            foreach (var uniteeTask in uniteeTaskList)
            {
                if(mergeUniteeTaskList.Count > 0)
                {
                    bool doublonFound = false;
                    foreach (var task in mergeUniteeTaskList)
                    {
                        if (uniteeTask.Date.Equals(task.Date) && uniteeTask.Project.Equals(task.Project) &&
                            uniteeTask.Description.Equals(task.Description))
                        {
                            task.TimeSpent = task.TimeSpent + uniteeTask.TimeSpent;
                            doublonFound = true;
                            break;
                        }   
                    }

                    if (!doublonFound)
                    {
                        mergeUniteeTaskList.Add(uniteeTask);
                    }
                }
                else
                {
                    mergeUniteeTaskList.Add(uniteeTask);
                }
            }


            return mergeUniteeTaskList;
        }

        private static void FillUniteeTaskList(List<UniteeTask> uniteeTaskList)
        {
            foreach (var line in File.ReadAllLines(path + inputFileName))
            {
                if (!(line.Contains("data-title=\"Day\"") && line.Contains("data-title=\"Task\"") && line.Contains("data-title=\"Level 1\"") && line.Contains("data-title=\"Time\"")))
                {
                    continue;
                }
                UniteeTask uniteeTask = new UniteeTask();
                string formattedLine = FormatLine(line);
                var split = formattedLine.Split(">");
                for (int i = 0; i < split.Length; i++)
                {
                    if (split[i].Contains("data-title=\"Day\""))
                    {
                        string literalDate = split[i + 1].Split("<")[0].Trim();
                        string[] s = literalDate.Split("-");
                        if (s.Length != 3)
                        {
                            break;
                        }
                        literalDate = s[2] + "/" + s[1] + "/" + s[0];
                        Console.Write(literalDate);
                        uniteeTask.Date = DateTime.ParseExact(literalDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    }

                    if (split[i].Contains("data-title=\"Task\""))
                    {
                        string description = split[i + 1].Split("<")[0].Trim();
                        uniteeTask.Description = description;
                        Console.Write(description);
                    }

                    if (split[i].Contains("data-title=\"Time\""))
                    {
                        string literalTime = split[i + 1].Split("<")[0].Trim().Replace(" ", "");
                        decimal nbMinutes = 0;
                        if (literalTime.Contains("h"))
                        {
                            nbMinutes += Decimal.Parse(literalTime.Split("h")[0].Trim()) * 60;
                            literalTime = literalTime.Substring(literalTime.IndexOf("h") + 1);
                        }
                        if (literalTime.Contains("m"))
                        {
                            nbMinutes += Decimal.Parse(literalTime.Split("m")[0].Trim());
                            literalTime = literalTime.Substring(literalTime.IndexOf("m") + 1);
                        }
                        if (literalTime.Contains("s"))
                        {
                            nbMinutes += Decimal.Parse(literalTime.Split("s")[0].Trim()) / 60;
                            literalTime = literalTime.Substring(literalTime.IndexOf("s") + 1);
                        }

                        uniteeTask.TimeSpent = nbMinutes;
                        Console.Write(nbMinutes);
                    }

                    if (split[i].Contains("data-title=\"Level 1\""))
                    {
                        string project = split[i + 1].Split("<")[0].Trim();
                        uniteeTask.Project = project;
                        Console.Write(project);
                    }
                }

                if (uniteeTask.Project != null && uniteeTask.Date != null && uniteeTask.Description != null)
                {
                    uniteeTaskList.Add(uniteeTask);
                    Console.WriteLine();
                }
            }
        }

        private static string FormatLine(string line)
        {
            var formattedLine = Regex.Replace(line, "<!--.*?-->", String.Empty, RegexOptions.Singleline);
            formattedLine = Regex.Replace(formattedLine, "<span.*?>", String.Empty, RegexOptions.Singleline);
            formattedLine = Regex.Replace(formattedLine, "<a.*?>", String.Empty, RegexOptions.Singleline);
            formattedLine = formattedLine.Replace("<div style=\" font-weight: bold;\">", "");
            return formattedLine;
        }
    }
}
