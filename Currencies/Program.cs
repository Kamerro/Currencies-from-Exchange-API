using System;
using System.IO;
using System.IO.Packaging;
using System.Net.Http;
//Async programming
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Win32.TaskScheduler;
//Install - Package EPPlus; open and edit excel file.
using OfficeOpenXml;
using System.IO;
//Into project needs to be installed Newtonsoft Json by PM console
using Newtonsoft.Json.Linq;
using static System.Threading.Thread;
using OfficeOpenXml.Drawing.Chart;
using System.Linq;

class Program
{
    //Creating static HttpClient 
    private static readonly HttpClient client = new HttpClient();

    static async System.Threading.Tasks.Task Main(string[] args)
    {
        //Define taskname and find the task in task scheduler.
        string taskName = "Open Notepad";
        Microsoft.Win32.TaskScheduler.Task existingTask = TaskService.Instance.FindTask(taskName);
        //Check if the task is already created
        if (existingTask == null)
        {
            CreateTask(taskName);
        }
        //Try to read data and save it to the file:
            try
            {
                //Creating HttpResponseMessage that waits for the client request
                HttpResponseMessage response = await client.GetAsync("https://api.exchangerate-api.com/v4/latest/USD");
                //Creating string represents the response content
                string responseBody = await response.Content.ReadAsStringAsync();
                //
                //Creating collection provided with NewtonsoftJson
                var data = JObject.Parse(responseBody);
                //Showing the actual currencies ;)
                Console.WriteLine("1 USD is equal to " + data["rates"]["EUR"] + " Euros");
                Console.WriteLine("1 USD is equal to " + data["rates"]["PLN"] + " PLN");
                using (StreamWriter st = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"Values.txt",append:true))
                {

                    st.WriteLine("1 USD is equal to " + data["rates"]["EUR"] + " Euros " + DateTime.Now.ToString());
                    st.WriteLine("1 USD is equal to " + data["rates"]["PLN"] + " PLN " + DateTime.Now.ToString());
                }

                //After saving data 
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Ustawienie kontekstu licencji
                using (ExcelPackage package = new ExcelPackage(new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\EValues.xlsx")))
                {
                ExcelWorksheet worksheet;
                if (package.Workbook.Worksheets.Count == 0)
                {
                    // Dodaj nowy arkusz jeżeli nie ma żadnego
                    worksheet = package.Workbook.Worksheets.Add("Arkusz1");
                }
                else
                {
                    // Użyj istniejącego arkusza
                    worksheet = package.Workbook.Worksheets[0];
                }

                //Add values to the sheet:
                int row = 1;
                while (worksheet.Cells[row, 1].Value != null && worksheet.Cells[row, 1].Value.ToString() != string.Empty)
                {
                    row++;
                }

                // Adds values to the first empty line:
                worksheet.Cells[row, 1].Value = "1 USD is equal to:";
                worksheet.Cells[row, 2].Value = data["rates"]["EUR"].ToObject<double>();
                worksheet.Cells[row, 3].Value = "Euros";
                worksheet.Cells[row, 4].Value = DateTime.Now.ToString();
                worksheet.Cells[row, 5].Value = "1 USD is equal to:";
                worksheet.Cells[row, 6].Value = data["rates"]["PLN"].ToObject<double>();
                worksheet.Cells[row, 7].Value = "PLN";


                // Szukamy wykresu o nazwie "Wykres"
                ExcelLineChart chart = null;
                if (worksheet.Drawings.Count > 0)
                {
                    chart = worksheet.Drawings.OfType<ExcelLineChart>().FirstOrDefault(c => c.Name == "Chart");
                }

                // Jeśli nie znaleziono wykresu, utwórz nowy
                if (chart == null)
                {
                    chart = (ExcelLineChart)worksheet.Drawings.AddChart("Chart", eChartType.Line);
                }

                // Usuń istniejące serie
                while (chart.Series.Count > 0) { chart.Series.Delete(0); }

                // Dodajemy dane do wykresu
                var series = chart.Series.Add(worksheet.Cells[1, 2, worksheet.Dimension.End.Row, 2], worksheet.Cells[1, 4, worksheet.Dimension.End.Row, 4]);
                series.Header = "USD to EUR";
                var series2 = chart.Series.Add(worksheet.Cells[1, 6, worksheet.Dimension.End.Row, 6], worksheet.Cells[1, 4, worksheet.Dimension.End.Row, 4]);
                series2.Header = "USD to PLN";
                // Ustawiamy pozycję i rozmiar wykresu
                chart.SetPosition(6, 0, 3, 0);
                chart.SetSize(800, 600);

                // Save the project:
                package.Save();
                }
        }
          catch { }
    }

    private static void CreateTask(string taskName)
    {
        TaskDefinition td = TaskService.Instance.NewTask();
        td.RegistrationInfo.Author = "Main";
        td.RegistrationInfo.Description = "Sample task opening Notepad";
        td.Actions.Add(new ExecAction(AppDomain.CurrentDomain.BaseDirectory+@"Currencies.exe"));
        TimeTrigger tt = new TimeTrigger();
        tt.Repetition.Interval = TimeSpan.FromHours(1); // Zadanie zostanie uruchomione co minute
        td.Triggers.Add(tt);
        TaskService.Instance.RootFolder.RegisterTaskDefinition(taskName, td);
        Console.WriteLine("Task scheduled.");
    }
}