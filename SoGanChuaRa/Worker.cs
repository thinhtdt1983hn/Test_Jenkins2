using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SoGanChuaRa
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);

                //Create COM Objects. Create a COM object for everything that is referenced
                Application xlApp = new Application();
                Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Downloads\Demo_data\Data_temp.xlsx");
                _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                //_Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                //int colCount = xlRange.Columns.Count;

                //====================================
                var _1G1_data = new List<int>();
                var _1DB_data = new List<int>();
                var _1B_data = new List<int>();

                //====================================

                for (int i = rowCount; i >= 1; i--)
                {
                    string _1G1_temp = xlRange.Cells[i, 1].Value2.ToString().Substring(3);
                    string _temp = xlRange.Cells[i, 2].Value2.ToString();

                    string _1B_temp = _temp.Substring(0, 2);
                    string _1DB_temp = string.Empty;

                    if (_temp.Length > 5) _1DB_temp = _temp.Substring(4);
                    else _1DB_temp = _temp.Substring(3);


                    if (!_1G1_data.Contains(Int32.Parse(_1G1_temp)) && _1G1_data.Count <= 100) _1G1_data.Add(Int32.Parse(_1G1_temp));

                    if (!_1DB_data.Contains(Int32.Parse(_1DB_temp)) && _1DB_data.Count <= 100) _1DB_data.Add(Int32.Parse(_1DB_temp));

                    if (!_1B_data.Contains(Int32.Parse(_1B_temp)) && _1B_data.Count <= 100) _1B_data.Add(Int32.Parse(_1B_temp));

                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);


                string _1G1_KQ = string.Empty;
                string _1DB_KQ = string.Empty;
                string _1B_KQ = string.Empty;

                for (int i = 0; i < _1G1_data.Count; i++)
                {
                    if (_1G1_data[i] < 10) _1G1_KQ += "0" + _1G1_data[i].ToString() + ",";
                    else _1G1_KQ += _1G1_data[i].ToString() + ",";
                }
                for (int i = 0; i < _1DB_data.Count; i++)
                {
                    if (_1DB_data[i] < 10) _1DB_KQ += "0" + _1DB_data[i].ToString() + ",";
                    else _1DB_KQ += _1DB_data[i].ToString() + ",";
                }
                for (int i = 0; i < _1B_data.Count; i++)
                {
                    if (_1B_data[i] < 10) _1B_KQ += "0" + _1B_data[i].ToString() + ",";
                    else _1B_KQ += _1B_data[i].ToString() + ",";
                }

                if (_1G1_data.Count < 100)
                {
                    for (int i = 0; i < 100; i++)
                    {
                        string _temp = i.ToString() + ",";
                        if (i < 10) _temp = "0" + i;

                        if (!_1G1_KQ.Contains(_temp))
                        {
                            _1G1_KQ += _temp;
                        }
                    }
                }
                if (_1DB_data.Count < 100)
                {
                    for (int i = 0; i < 100; i++)
                    {
                        string _temp = i.ToString() + ",";
                        if (i < 10) _temp = "0" + i;

                        if (!_1DB_KQ.Contains(_temp))
                        {
                            _1DB_KQ += _temp;
                        }
                    }
                }
                if (_1B_data.Count < 100)
                {
                    for (int i = 0; i < 100; i++)
                    {
                        string _temp = i.ToString() + ",";
                        if (i < 10) _temp = "0" + i;

                        if (!_1B_KQ.Contains(_temp))
                        {
                            _1B_KQ += _temp;
                        }
                    }
                }

                string fileName = @"D:\Downloads\Demo_data\Data_temp_export.txt";

                try
                {
                    // Check if file already exists. If yes, delete it.
                    if (File.Exists(fileName))
                    {
                        File.Delete(fileName);
                    }

                    using (StreamWriter sw = File.CreateText(fileName))
                    {
                        // 1G1
                        sw.WriteLine("1G1 So Gan: {0} so tu moi nhat", _1G1_data.Count());
                        sw.WriteLine(_1G1_KQ);
                        sw.WriteLine(" ");
                        sw.WriteLine(" ");

                        // 1DB
                        sw.WriteLine("1DB_So Gan: {0} so tu moi nhat", _1DB_data.Count());
                        sw.WriteLine(_1DB_KQ);
                        sw.WriteLine(" ");
                        sw.WriteLine(" ");

                        // 1B
                        sw.WriteLine("1B_So Gan: {0} so tu moi nhat", _1B_data.Count());
                        sw.WriteLine(_1B_KQ);
                    }

                }
                catch (Exception Ex)
                {
                    Console.WriteLine(Ex.ToString());
                }


                _logger.LogInformation("Worker running DONE: {time}", DateTimeOffset.Now);
                await Task.Delay(10000, stoppingToken);
            }
        }
    }
}