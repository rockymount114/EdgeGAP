using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Threading.Tasks;

namespace EdgeGAP
{
    public partial class Ribbon1
    {
        private System.Windows.Forms.Timer statusUpdateTimer;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            statusUpdateTimer = new System.Windows.Forms.Timer();
            statusUpdateTimer.Interval = 100; // Update every 100ms
            statusUpdateTimer.Tick += StatusUpdateTimer_Tick;
        }

        private string currentStatus = "";

        private void StatusUpdateTimer_Tick(object sender, EventArgs e)
        {
            UpdateStatus(currentStatus);
        }

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = folderDialog.SelectedPath;
                    statusUpdateTimer.Start();
                    await ConvertTxtToCsvAsync(folderPath);
                    statusUpdateTimer.Stop();
                }
            }
        }

        private async Task ConvertTxtToCsvAsync(string folderPath)
        {
            var columnWidths = new int[] { 2, 5, 12, 25, 4, 3, 2, 4, 6, 6, 15, 8, 8, 25, 20, 20, 3, 3, 1, 25, 20, 20, 3, 3, 1, 25, 25, 20, 2, 9, 25, 25, 20, 2, 9, 25, 25, 20, 2, 9, 8, 4, 8, 8, 6, 3, 3, 1, 3, 7, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 9, 15, 25, 20, 20, 3, 4, 25, 25, 20, 2, 9, 25, 20, 20, 3, 4, 25, 25, 20, 2, 9, 25, 20, 20, 3, 4, 25, 25, 20, 2, 9, 9, 1, 9, 4, 1, 1, 10, 1, 1, 2, 29, 6, 54, 10, 10, 10 };

            var columnHeaders = new string[] { "N", "County", "AccountNumber", "VIN", "Brand", "Col3", "Style", "Year", "C1", "C2", "ParcelNo", "PurchasedDate", "C3", "FristName", "MiddleName", "LastName", "Title", "C4", "C5", "FirstName2", "MiddleName2", "LastName2", "C33", "C43", "C41", "25zx", "25xc", "20cv", "2vb", "9bn", "Address1", "Apt", "City", "ST", "ZipCode", "Address2", "A25", "City2", "ST2", "ZipCode2", "RegNumber", "4", "YearFrom", "YearEnd", "DueValue", "3sd", "3df", "1fg", "3hj", "Account?", "1a", "1b", "1c", "1d", "1f", "1g", "1h", "1l", "1m", "1n", "1o", "AccountNo?", "15a", "25b", "20c", "20d", "3ee", "4dd", "25e", "25f", "20g", "2hh", "9ii", "25jj", "20k", "20kk", "3ll", "4mm", "25oo", "2pp5", "20qq", "2rr", "9ss", "FirstName3", "MiddleName3", "LastName3", "3tt", "4qw", "Add3", "Add3Line2", "City3", "ST3", "ZipCode3", "9as", "1er", "AccountNumber?", "CarType", "1yu", "1gh", "10?", "cv", "bn", "2?", "mj", "Acct?", "54", "P1", "P2", "P3" };

            string[] txtFiles = Directory.GetFiles(folderPath, "*.txt");
            int totalFiles = txtFiles.Length;
            int processedFiles = 0;

            var allData = new List<List<string>>();

            await Task.Run(() =>
            {
                Parallel.ForEach(txtFiles, filePath =>
                {
                    var fileData = ReadFixedWidthFile(filePath, columnWidths);
                    lock (allData)
                    {
                        allData.AddRange(fileData);
                    }

                    // Save individual CSV file
                    string csvFilePath = Path.ChangeExtension(filePath, ".csv");
                    WriteCsvFile(csvFilePath, fileData, columnHeaders);

                    int currentProcessed = System.Threading.Interlocked.Increment(ref processedFiles);
                    int percentage = (currentProcessed * 100) / totalFiles;
                    currentStatus = $"Progress: {percentage}% - Processed {currentProcessed} of {totalFiles} files";
                });
            });

            // Display data in Excel
            currentStatus = "Displaying data in Excel...";
            await DisplayDataInExcelAsync(allData, columnHeaders);

            currentStatus = "Conversion completed successfully!";
            await Task.Delay(2000); // Display the completion message for 2 seconds
            Globals.ThisAddIn.Application.StatusBar = false; // Clear the status bar
        }

        private void UpdateStatus(string message)
        {
            Globals.ThisAddIn.Application.StatusBar = message;
        }

        private List<List<string>> ReadFixedWidthFile(string filePath, int[] columnWidths)
        {
            var data = new List<List<string>>();
            using (var reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var row = new List<string>(columnWidths.Length);
                    int startIndex = 0;
                    foreach (int width in columnWidths)
                    {
                        if (startIndex + width <= line.Length)
                        {
                            row.Add(line.Substring(startIndex, width).Trim());
                        }
                        else
                        {
                            row.Add(string.Empty);
                        }
                        startIndex += width;
                    }
                    data.Add(row);
                }
            }
            return data;
        }

        private void WriteCsvFile(string filePath, List<List<string>> data, string[] headers)
        {
            using (var writer = new StreamWriter(filePath))
            {
                writer.WriteLine(string.Join(",", headers));
                foreach (var row in data)
                {
                    writer.WriteLine(string.Join(",", row.Select(field => $"\"{field.Replace("\"", "\"\"")}\"")));
                }
            }
        }

        private async Task DisplayDataInExcelAsync(List<List<string>> data, string[] headers)
        {
            var excel = Globals.ThisAddIn.Application;
            var workbook = excel.ActiveWorkbook;
            var worksheet = workbook.Worksheets[1] as Excel.Worksheet;
            worksheet.Cells.Clear();

            // Write headers
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, headers.Length]].Value = headers;

            // Prepare data array
            object[,] dataArray = new object[data.Count, headers.Length];

            await Task.Run(() =>
            {
                for (int row = 0; row < data.Count; row++)
                {
                    for (int col = 0; col < headers.Length; col++)
                    {
                        dataArray[row, col] = col < data[row].Count ? data[row][col] : string.Empty;
                    }

                    if (row % 100 == 0) // Update status every 100 rows
                    {
                        int percentage = (row * 100) / data.Count;
                        currentStatus = $"Preparing data: {percentage}% - Row {row} of {data.Count}";
                    }
                }
            });

            // Write data to Excel in one operation
            currentStatus = "Writing data to Excel...";
            worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[data.Count + 1, headers.Length]].Value = dataArray;

            // Auto-fit columns
            worksheet.Columns.AutoFit();
        }
    }
}