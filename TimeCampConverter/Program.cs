using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace TimecampConverter
{
    class Program
    {
        private const string ApiToken = "9b704c896acbe4eb55e617dc5a";
        private const string ReferenceFacture = "2019-3";
        private static string path = "C:\\Users\\coren\\Desktop\\TimeCamp\\";
        private static string outputFileName = "Export_Mars.xlsx";
        private static readonly DateTime start = new DateTime(2019,03,01);
        private static readonly DateTime end = new DateTime(2019,03,31);

        static void Main(string[] args)
        {
            Console.WriteLine("TIMECAMP CONVERT");
            Console.WriteLine("Reading File");
            //TODO controller les max (En heure (max 50h/s ou 140h/m))
            List<UniteeTask> uniteeTaskList =  FillUniteeTaskList(start,end).Result;
            //Suppression des doublons
            uniteeTaskList = MergeDoubleTaskByDay(uniteeTaskList);
            uniteeTaskList.Sort(new UniteeTaskComparator());

            //Création du fichier Excel
            // create xls if not exists
            if (!File.Exists(path+outputFileName))
            {
                XSSFWorkbook wb;
                XSSFSheet sh;
                wb = new XSSFWorkbook();

                // create sheet
                sh = (XSSFSheet)wb.CreateSheet("Sheet1");
                // Create header
                var r = sh.CreateRow(0);
                //Date 
                r.CreateCell(0).SetCellValue("Date");
                //Temps facturé 
                r.CreateCell(1).SetCellValue("Temps facturé");
                //Projet
                r.CreateCell(2).SetCellValue("Projet");
                //Type    
                r.CreateCell(3).SetCellValue("Type");
                //Réf facture
                r.CreateCell(4).SetCellValue("Référence facture");
                //Description
                r.CreateCell(5).SetCellValue("Description");

                int i = 1;
                for (i = 1; i < uniteeTaskList.Count; i++)
                {
                    var task = uniteeTaskList[i];
                    r = sh.CreateRow(i);
                    //Date 
                    r.CreateCell(0).SetCellValue(task.Date.ToString());
                    //Temps facturé 
                    r.CreateCell(1).SetCellValue((double)(task.TimeSpent / 60));
                    //Projet
                    r.CreateCell(2).SetCellValue(task.Project);
                    //Type    
                    r.CreateCell(3).SetCellValue("Developpement");
                    //Réf facture
                    r.CreateCell(4).SetCellValue(task.ReferenceFacture);
                    //Description
                    r.CreateCell(5).SetCellValue(task.Description);
                }

                var limitTableau = ++i;
                r = sh.CreateRow(i++);
                //Date 
                r.CreateCell(0).SetCellValue("Total");
                //Temps facturé 
                r.CreateCell(1).SetCellFormula("SUBTOTAL(109,B2:B" + limitTableau + ")");
                
                r = sh.CreateRow(i++);
                //Date 
                r.CreateCell(0).SetCellValue("Total");
                //Temps facturé 
                r.CreateCell(1).SetCellFormula("SUBTOTAL(109,B2:B" + limitTableau + ")/8/2");

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

        private static async Task<List<UniteeTask>> FillUniteeTaskList(DateTime start , DateTime end)
        {
            HttpClient httpClient = new HttpClient();
            var serializer = new DataContractJsonSerializer(typeof(List<Entrie>));
            var response = httpClient.GetStreamAsync("https://www.timecamp.com/third_party/api/entries/format/json/api_token/" + ApiToken + "/from/" + start.ToString("yyyy-MM-dd") + "/to/" + end.ToString("yyyy-MM-dd") + "/");
            var entries = serializer.ReadObject(await response) as List<Entrie>;
            List<UniteeTask> uniteeTaskList = new List<UniteeTask>();
            foreach (var entrie in entries)
            {
                UniteeTask uniteeTask = new UniteeTask();
                uniteeTask.Date = DateTime.ParseExact(entrie.date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                DateTime endTime = uniteeTask.Date
                    .AddHours(Convert.ToDouble(entrie.end_time.Split(":")[0]))
                    .AddMinutes(Convert.ToDouble(entrie.end_time.Split(":")[1]))
                    .AddSeconds(Convert.ToDouble(entrie.end_time.Split(":")[2]));
                DateTime startTime = uniteeTask.Date
                    .AddHours(Convert.ToDouble(entrie.start_time.Split(":")[0]))
                    .AddMinutes(Convert.ToDouble(entrie.start_time.Split(":")[1]))
                    .AddSeconds(Convert.ToDouble(entrie.start_time.Split(":")[2]));
                uniteeTask.TimeSpent = (decimal) endTime.Subtract(startTime).TotalMinutes;

                // Get task
                serializer = new DataContractJsonSerializer(typeof(TimeCampTask));
                response = httpClient.GetStreamAsync("https://www.timecamp.com/third_party/api/tasks/format/json/api_token/" + ApiToken + "/task_id/" + entrie.task_id);
                var task = serializer.ReadObject(await response) as TimeCampTask;
                uniteeTask.Description = task.name;
                
                // Get project tack
                serializer = new DataContractJsonSerializer(typeof(TimeCampTask));
                response = httpClient.GetStreamAsync("https://www.timecamp.com/third_party/api/tasks/format/json/api_token/" + ApiToken + "/task_id/" + task.parent_id);
                var projectTask = serializer.ReadObject(await response) as TimeCampTask;
                uniteeTask.Project = projectTask.name;

                uniteeTask.ReferenceFacture = ReferenceFacture;


                if (uniteeTask.Project != null && uniteeTask.Date != null && uniteeTask.Description != null)
                {
                    uniteeTaskList.Add(uniteeTask);
                    Console.WriteLine();
                }
            }
            return uniteeTaskList;
        }

        
    }
}
